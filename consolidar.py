"""Consolida relatórios operacionais .xlsx de várias equipes.

Uso: python consolidar.py [--entrada PASTA] [--saida PASTA]
"""

from __future__ import annotations

import argparse
import os
import re
import tempfile
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.table import Table, TableStyleInfo


BASE = Path(__file__).resolve().parent
COLUNAS = [
    "data", "ordem_servico", "equipe", "servico", "local", "colaborador",
    "status", "meta", "realizado", "horas_trabalhadas", "horas_extras",
]
NOMES = {
    "data": "Data", "ordem_servico": "Ordem de serviço", "equipe": "Equipe",
    "servico": "Serviço", "local": "Local", "colaborador": "Colaborador",
    "status": "Status", "meta": "Meta", "realizado": "Realizado",
    "horas_trabalhadas": "Horas trabalhadas", "horas_extras": "Horas extras",
}
ALIASES = {
    "data": ("data", "data execucao", "data do servico"),
    "ordem_servico": ("ordem de servico", "ordem servico", "os", "id os", "codigo os"),
    "equipe": ("equipe", "time"),
    "servico": ("servico", "tipo de servico", "atividade"),
    "local": ("local", "unidade", "filial"),
    "colaborador": ("colaborador", "funcionario", "responsavel"),
    "status": ("status", "situacao"),
    "meta": ("meta", "meta prevista"),
    "realizado": ("realizado", "producao", "quantidade realizada"),
    "horas_trabalhadas": ("horas trabalhadas", "horas", "horas normais"),
    "horas_extras": ("horas extras", "h extras", "hora extra"),
}
SERVICOS = {
    "manutencao": "Manutenção", "inspecao": "Inspeção",
    "instalacao": "Instalação", "atendimento": "Atendimento",
}
STATUS = {
    "concluido": "Concluído", "concluida": "Concluído", "finalizado": "Concluído",
    "finalizada": "Concluído", "em andamento": "Em andamento", "andamento": "Em andamento",
    "aberto": "Em andamento", "pendente": "Pendente", "nao iniciado": "Pendente",
    "cancelado": "Cancelado", "cancelada": "Cancelado",
}
AZUL = "142C3F"
VERDE = "218A7C"
CLARO = "EAF3F5"
BRANCO = "FFFFFF"
LARANJA = "BD7138"


def chave(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", texto.lower())).strip()


MAPA_CABECALHOS = {chave(alias): campo for campo, aliases in ALIASES.items() for alias in aliases}


def texto_limpo(valor: Any) -> str:
    return re.sub(r"\s+", " ", str(valor) if valor is not None else "").strip()


def ler_data(valor: Any) -> date:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and 1 <= valor <= 2958465:
        return from_excel(valor).date()
    texto = texto_limpo(valor)
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    raise ValueError("data inválida; use DD/MM/AAAA ou AAAA-MM-DD")


def ler_numero(valor: Any) -> float:
    if isinstance(valor, bool):
        raise ValueError("número inválido")
    if isinstance(valor, (int, float)):
        numero = float(valor)
    else:
        texto = texto_limpo(valor).replace(" ", "").replace("R$", "")
        if not texto:
            raise ValueError("valor obrigatório")
        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")
        try:
            numero = float(texto)
        except ValueError as exc:
            raise ValueError("número inválido") from exc
    if not (-1e12 < numero < 1e12):
        raise ValueError("número fora do intervalo permitido")
    return round(numero, 2)


@dataclass
class Resultado:
    arquivos_lidos: int
    linhas_lidas: int
    linhas_validas: int
    linhas_descartadas: int
    inconsistencias: int
    pasta_saida: Path


def _estilizar_cabecalho(ws, linha: int, total_colunas: int) -> None:
    ws.row_dimensions[linha].height = 27
    for celula in ws[linha][:total_colunas]:
        celula.fill = PatternFill("solid", fgColor=AZUL)
        celula.font = Font(name="Aptos", size=10, bold=True, color=BRANCO)
        celula.alignment = Alignment(vertical="center")


def _tabela(ws, linha_cabecalho: int, ultima_linha: int, nome: str) -> None:
    ws.freeze_panes = f"A{linha_cabecalho + 1}"
    ws.sheet_view.showGridLines = False
    if ultima_linha > linha_cabecalho:
        tabela = Table(displayName=nome, ref=f"A{linha_cabecalho}:{get_column_letter(ws.max_column)}{ultima_linha}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)
    else:
        ws.auto_filter.ref = f"A{linha_cabecalho}:{get_column_letter(ws.max_column)}{linha_cabecalho}"


def _larguras(ws, larguras: list[int]) -> None:
    for indice, largura in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(indice)].width = largura


def _salvar(wb: Workbook, destino: Path) -> None:
    with tempfile.NamedTemporaryFile(dir=destino.parent, suffix=".xlsx", delete=False) as arquivo:
        temporario = Path(arquivo.name)
    try:
        wb.save(temporario)
        os.replace(temporario, destino)
    finally:
        temporario.unlink(missing_ok=True)


def _planilha_consolidada(registros: list[dict], destino: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados consolidados"
    ws.append([NOMES[c] for c in COLUNAS] + ["Arquivo origem", "Aba origem", "Linha origem"])
    for registro in registros:
        ws.append([registro[c] for c in COLUNAS] + [registro["arquivo"], registro["aba"], registro["linha"]])
    _estilizar_cabecalho(ws, 1, 14)
    _tabela(ws, 1, ws.max_row, "DadosOperacionais")
    _larguras(ws, [15, 21, 16, 19, 18, 23, 19, 14, 14, 21, 16, 20, 20, 15])
    for linha in ws.iter_rows(min_row=2):
        linha[0].number_format = "dd/mm/yyyy"
        for indice in (7, 8, 9, 10):
            linha[indice].number_format = "#,##0.00"
    _salvar(wb, destino)


def _planilha_inconsistencias(problemas: list[dict], destino: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Inconsistências"
    campos = ["Arquivo", "Aba", "Linha", "Ordem de serviço", "Campo", "Problema", "Valor original", "Ação"]
    ws.append(campos)
    for p in problemas:
        ws.append([p.get(c, "") for c in campos])
    _estilizar_cabecalho(ws, 1, len(campos))
    _tabela(ws, 1, ws.max_row, "ProblemasEncontrados")
    _larguras(ws, [23, 22, 12, 23, 22, 51, 27, 25])
    for linha in ws.iter_rows(min_row=2):
        linha[5].alignment = Alignment(wrap_text=True, vertical="center")
        linha[7].font = Font(color=LARANJA)
        ws.row_dimensions[linha[0].row].height = 31
    _salvar(wb, destino)


def _grupo(registros: list[dict], campo: str) -> list[tuple]:
    totais: dict[Any, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for registro in registros:
        grupo = registro[campo]
        totais[grupo]["Registros"] += 1
        for nome in ("meta", "realizado", "horas_trabalhadas", "horas_extras"):
            totais[grupo][nome] += registro[nome]
    linhas = []
    for grupo, v in sorted(totais.items()):
        meta, realizado, horas = v["meta"], v["realizado"], v["horas_trabalhadas"]
        linhas.append((grupo, int(v["Registros"]), meta, realizado,
                      realizado / meta if meta else 0, horas,
                      realizado / horas if horas else 0, v["horas_extras"]))
    return linhas


def _aba_grupo(wb: Workbook, titulo: str, descricao: str, linhas: list[tuple], grafico: str) -> None:
    ws = wb.create_sheet(titulo)
    ws.merge_cells("A1:H1")
    ws["A1"] = descricao
    ws["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=BRANCO)
    ws["A1"].fill = PatternFill("solid", fgColor=AZUL)
    ws.row_dimensions[1].height = 36
    ws["A2"] = "Dados fictícios • linhas válidas consolidadas"
    ws["A2"].font = Font(name="Aptos", size=10, color=VERDE)
    ws.append([])
    ws.append([titulo.replace("Por ", ""), "Registros", "Meta", "Realizado", "% cumprimento",
               "Horas trabalhadas", "Produtividade (unid./h)", "Horas extras"])
    for linha in linhas:
        ws.append(linha)
    _estilizar_cabecalho(ws, 4, 8)
    _tabela(ws, 4, ws.max_row, "Resumo" + chave(titulo).replace(" ", "").title())
    _larguras(ws, [23, 14, 17, 17, 20, 23, 27, 19])
    for linha in ws.iter_rows(min_row=5):
        linha[4].number_format = "0.0%"
        for indice in (2, 3, 5, 6, 7):
            linha[indice].number_format = "#,##0.00"
    if linhas:
        chart = LineChart() if grafico == "linha" else BarChart()
        chart.title = "Meta x realizado"
        chart.y_axis.title = "Unidades"
        chart.x_axis.title = titulo.replace("Por ", "")
        chart.height, chart.width = 8, 16
        chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=4, max_row=ws.max_row), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=5, max_row=ws.max_row))
        ws.add_chart(chart, "A" + str(ws.max_row + 3))


def _planilha_resumo(registros: list[dict], resultado: Resultado, destino: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Visão geral"
    ws.merge_cells("A1:D1")
    ws["A1"] = "RESUMO OPERACIONAL"
    ws["A1"].font = Font(name="Aptos Display", size=20, bold=True, color=BRANCO)
    ws["A1"].fill = PatternFill("solid", fgColor=AZUL)
    ws.row_dimensions[1].height = 40
    ws["A2"] = "Consolidação automática • dados fictícios"
    ws["A2"].font = Font(name="Aptos", size=10, color=VERDE)
    meta = sum(r["meta"] for r in registros)
    realizado = sum(r["realizado"] for r in registros)
    horas = sum(r["horas_trabalhadas"] for r in registros)
    extras = sum(r["horas_extras"] for r in registros)
    indicadores = [
        ("Arquivos lidos", resultado.arquivos_lidos),
        ("Linhas analisadas", resultado.linhas_lidas),
        ("Registros válidos", resultado.linhas_validas),
        ("Linhas descartadas", resultado.linhas_descartadas),
        ("Inconsistências", resultado.inconsistencias),
        ("Meta total", meta), ("Realizado total", realizado),
        ("Cumprimento", realizado / meta if meta else 0),
        ("Horas trabalhadas", horas),
        ("Produtividade (unid./h)", realizado / horas if horas else 0),
        ("Horas extras", extras),
    ]
    for indice, (nome, valor) in enumerate(indicadores, 4):
        ws.cell(indice, 1, nome)
        ws.cell(indice, 2, valor)
        ws.cell(indice, 1).fill = PatternFill("solid", fgColor=CLARO)
        ws.cell(indice, 1).font = Font(name="Aptos", bold=True, color=AZUL)
        ws.cell(indice, 2).font = Font(name="Aptos Display", size=13, bold=True, color=VERDE)
        ws.row_dimensions[indice].height = 25
    for linha in (9, 10, 12, 13, 14):
        ws.cell(linha, 2).number_format = "#,##0.00"
    ws["B11"].number_format = "0.0%"
    ws["A16"] = "Regra: linhas com campos obrigatórios inválidos ou OS duplicada são descartadas."
    ws["A17"] = "Consulte inconsistencias.xlsx para corrigir a origem e execute novamente."
    for linha in (16, 17):
        ws.cell(linha, 1).font = Font(name="Aptos", size=10, color=AZUL)
    ws.sheet_view.showGridLines = False
    _larguras(ws, [39, 24, 18, 18])
    for campo, titulo, descricao, grafico in (
        ("equipe", "Por equipe", "DESEMPENHO POR EQUIPE", "barra"),
        ("servico", "Por serviço", "DESEMPENHO POR SERVIÇO", "barra"),
        ("mes", "Por mês", "EVOLUÇÃO MENSAL", "linha"),
    ):
        _aba_grupo(wb, titulo, descricao, _grupo(registros, campo), grafico)
    _salvar(wb, destino)


def consolidar_pasta(entrada: Path, saida: Path) -> Resultado:
    entrada, saida = Path(entrada).resolve(), Path(saida).resolve()
    if not entrada.is_dir():
        raise ValueError(f"Pasta de entrada não encontrada: {entrada}")
    arquivos = sorted(p for p in entrada.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$"))
    if not arquivos:
        raise ValueError(f"Nenhum arquivo .xlsx encontrado em {entrada}")
    registros: list[dict] = []
    problemas: list[dict] = []
    vistos: set[str] = set()
    linhas_lidas = linhas_descartadas = arquivos_lidos = 0

    def registrar(arquivo: str, aba: str, linha: int | str, os_id: str, campo: str,
                 mensagem: str, valor: Any, acao: str = "Linha descartada") -> None:
        problemas.append({"Arquivo": arquivo, "Aba": aba, "Linha": linha,
                          "Ordem de serviço": os_id, "Campo": campo,
                          "Problema": mensagem, "Valor original": texto_limpo(valor), "Ação": acao})

    for arquivo in arquivos:
        try:
            wb = load_workbook(arquivo, read_only=True, data_only=True)
        except Exception as exc:
            registrar(arquivo.name, "", "", "", "Arquivo", f"Não foi possível abrir: {exc}", "", "Arquivo ignorado")
            continue
        arquivos_lidos += 1
        try:
            for ws in wb.worksheets:
                if ws.max_row == 0:
                    continue
                cabecalho = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
                if not any(cabecalho):
                    continue
                indices = {}
                for indice, nome in enumerate(cabecalho):
                    campo = MAPA_CABECALHOS.get(chave(nome))
                    if campo and campo not in indices:
                        indices[campo] = indice
                faltantes = [NOMES[c] for c in COLUNAS if c not in indices]
                if faltantes:
                    registrar(arquivo.name, ws.title, 1, "", "Cabeçalho",
                             "Colunas obrigatórias ausentes: " + ", ".join(faltantes),
                             ", ".join(texto_limpo(v) for v in cabecalho), "Aba ignorada")
                    continue
                for numero_linha, celulas in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
                    if not any(v is not None and texto_limpo(v) for v in celulas):
                        continue
                    linhas_lidas += 1
                    bruto = {campo: celulas[idx] if idx < len(celulas) else None for campo, idx in indices.items()}
                    os_id = re.sub(r"\s+", "", texto_limpo(bruto["ordem_servico"])).upper()
                    erros: list[tuple[str, str, Any]] = []
                    registro: dict[str, Any] = {"ordem_servico": os_id}
                    if not os_id:
                        erros.append(("Ordem de serviço", "valor obrigatório", bruto["ordem_servico"]))
                    elif os_id in vistos:
                        erros.append(("Ordem de serviço", "OS duplicada em outro registro", bruto["ordem_servico"]))
                    try:
                        registro["data"] = ler_data(bruto["data"])
                    except ValueError as exc:
                        erros.append(("Data", str(exc), bruto["data"]))
                    for campo in ("equipe", "servico", "local", "colaborador"):
                        valor = texto_limpo(bruto[campo])
                        if not valor:
                            erros.append((NOMES[campo], "valor obrigatório", bruto[campo]))
                        else:
                            normalizado = chave(valor)
                            if campo == "servico":
                                valor = SERVICOS.get(normalizado, valor.title())
                            elif campo == "equipe":
                                valor = valor.title()
                            elif campo == "local":
                                valor = valor.title()
                            registro[campo] = valor
                    status = STATUS.get(chave(bruto["status"]))
                    if status is None:
                        erros.append(("Status", "status não reconhecido", bruto["status"]))
                    else:
                        registro["status"] = status
                    for campo in ("meta", "realizado", "horas_trabalhadas", "horas_extras"):
                        try:
                            numero = ler_numero(bruto[campo])
                            if numero < 0 or (campo == "horas_trabalhadas" and numero == 0):
                                raise ValueError("deve ser positivo" if campo == "horas_trabalhadas" else "não pode ser negativo")
                            registro[campo] = numero
                        except ValueError as exc:
                            erros.append((NOMES[campo], str(exc), bruto[campo]))
                    if erros:
                        linhas_descartadas += 1
                        for campo, mensagem, valor in erros:
                            registrar(arquivo.name, ws.title, numero_linha, os_id, campo, mensagem, valor)
                        continue
                    vistos.add(os_id)
                    registro.update(arquivo=arquivo.name, aba=ws.title, linha=numero_linha,
                                    mes=registro["data"].strftime("%Y-%m"))
                    registros.append(registro)
        finally:
            wb.close()

    registros.sort(key=lambda r: (r["data"], r["equipe"], r["ordem_servico"]))
    resultado = Resultado(arquivos_lidos, linhas_lidas, len(registros), linhas_descartadas, len(problemas), saida)
    saida.mkdir(parents=True, exist_ok=True)
    _planilha_consolidada(registros, saida / "consolidado.xlsx")
    _planilha_inconsistencias(problemas, saida / "inconsistencias.xlsx")
    _planilha_resumo(registros, resultado, saida / "resumo.xlsx")
    return resultado


def main() -> None:
    parser = argparse.ArgumentParser(description="Consolida planilhas operacionais de equipes.")
    parser.add_argument("--entrada", type=Path, default=BASE / "relatorios", help="Pasta com arquivos .xlsx")
    parser.add_argument("--saida", type=Path, default=BASE / "resultado", help="Pasta para os três arquivos gerados")
    args = parser.parse_args()
    try:
        resultado = consolidar_pasta(args.entrada, args.saida)
    except (ValueError, OSError) as exc:
        parser.exit(1, f"Erro: {exc}\n")
    print(f"Arquivos lidos: {resultado.arquivos_lidos}")
    print(f"Linhas: {resultado.linhas_validas} válidas, {resultado.linhas_descartadas} descartadas")
    print(f"Inconsistências: {resultado.inconsistencias}")
    print(f"Arquivos gerados em: {resultado.pasta_saida}")


if __name__ == "__main__":
    main()
