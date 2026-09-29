"""Cria quatro planilhas fictícias e reproduzíveis para a demonstração."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


BASE = Path(__file__).resolve().parent
DESTINO = BASE / "relatorios"
SERVICOS = ["Manutenção", "Inspeção", "Instalação", "Atendimento"]
CABECALHOS = {
    "a": ["Data", "Ordem de serviço", "Equipe", "Serviço", "Local", "Colaborador",
          "Status", "Meta", "Realizado", "Horas trabalhadas", "Horas extras"],
    "b": ["Data execução", "OS", "Time", "Tipo de serviço", "Unidade", "Funcionário",
          "Situação", "Meta prevista", "Produção", "Horas", "H. extras"],
    "c": ["Data", "Ordem de Serviço", "Equipe", "Atividade", "Filial", "Responsável",
          "Status", "Meta", "Quantidade realizada", "Horas normais", "Hora extra"],
    "d": ["Data do serviço", "Código OS", "Equipe", "Serviço", "Local", "Colaborador",
          "Status", "Meta", "Realizado", "Horas trabalhadas", "Horas extras"],
}


def criar() -> None:
    DESTINO.mkdir(parents=True, exist_ok=True)
    for indice_equipe, letra in enumerate("abcd"):
        wb = Workbook()
        ws = wb.active
        ws.title = f"Equipe {letra.upper()}"
        ws.append(CABECALHOS[letra])
        for mes in (7, 8, 9):
            for numero in range(1, 7):
                dia = 3 + numero * 3
                quando = date(2026, mes, dia)
                os_id = f"OS-{letra.upper()}-{mes:02d}-{numero:03d}"
                servico = SERVICOS[(indice_equipe + numero + mes) % len(SERVICOS)]
                meta = 80 + indice_equipe * 12 + numero * 4
                status = "Concluído" if numero <= 4 else "Em andamento" if numero == 5 else "Pendente"
                realizado = round(meta * (0.82 + ((mes + numero + indice_equipe) % 5) * 0.035), 1) if numero < 6 else 0
                horas = 7.5 + (numero % 3) * 0.5
                extras = 1.5 if numero in (3, 4) else 0
                colaborador = f"Funcionário {indice_equipe * 6 + numero:02d}"
                local = f"Unidade {'ABC'[(indice_equipe + numero) % 3]}"
                if letra == "a":
                    linha = [quando, os_id, "Equipe A", servico, local, colaborador,
                             status, meta, realizado, horas, extras]
                elif letra == "b":
                    linha = [quando.strftime("%d/%m/%Y"), os_id.lower(), "equipe b", servico.lower(),
                             local.lower(), colaborador, "finalizado" if numero <= 4 else status.lower(),
                             str(meta).replace(".", ","), str(realizado).replace(".", ","),
                             str(horas).replace(".", ","), str(extras).replace(".", ",")]
                elif letra == "c":
                    linha = [quando.strftime("%Y-%m-%d"), os_id, "EQUIPE C", servico,
                             local, colaborador, status, meta, realizado, horas, extras]
                else:
                    linha = [quando.strftime("%d-%m-%Y"), os_id, "Equipe D", servico,
                             local, colaborador, status, meta, realizado, horas, extras]

                # Cinco erros de origem demonstram as validações do processamento.
                if letra == "b" and mes == 8 and numero == 2:
                    linha[0] = "31/02/2026"  # Data impossível.
                if letra == "c" and mes == 9 and numero == 3:
                    linha[8] = -5  # Realizado negativo.
                if letra == "c" and mes == 8 and numero == 1:
                    linha[9] = "oito"  # Horas fora do formato numérico.
                if letra == "d" and mes == 7 and numero == 4:
                    linha[5] = None  # Colaborador ausente.
                if letra == "d" and mes == 9 and numero == 5:
                    linha[1] = "OS-A-07-001"  # OS já registrada na equipe A.
                ws.append(linha)

        for celula in ws[1]:
            celula.fill = PatternFill("solid", fgColor="142C3F")
            celula.font = Font(name="Aptos", color="FFFFFF", bold=True)
            celula.alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 26
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        for coluna, largura in enumerate((17, 23, 17, 21, 18, 23, 19, 13, 16, 22, 17), 1):
            ws.column_dimensions[get_column_letter(coluna)].width = largura
        if letra == "a":
            for linha in ws.iter_rows(min_row=2):
                linha[0].number_format = "dd/mm/yyyy"
        tabela = Table(displayName=f"RelatorioEquipe{letra.upper()}", ref=f"A1:K{ws.max_row}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)
        wb.save(DESTINO / f"equipe_{letra}.xlsx")


if __name__ == "__main__":
    criar()
    print(f"4 planilhas fictícias criadas em {DESTINO}")
