"""Teste de integração: entradas heterogêneas, rejeições e saídas finais."""

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from consolidar import consolidar_pasta


class ConsolidacaoTest(unittest.TestCase):
    def test_consolida_formatos_e_registra_inconsistencias(self):
        with tempfile.TemporaryDirectory() as temporario:
            raiz = Path(temporario)
            entrada, saida = raiz / "relatorios", raiz / "resultado"
            entrada.mkdir()

            wb_a = Workbook()
            ws_a = wb_a.active
            ws_a.append(["Data", "Ordem de serviço", "Equipe", "Serviço", "Local",
                         "Colaborador", "Status", "Meta", "Realizado", "Horas trabalhadas", "Horas extras"])
            ws_a.append(["01/09/2026", "OS-001", "Equipe A", "Manutenção", "Unidade A",
                         "Ana", "Concluído", 100, 80, 8, 0])
            ws_a.append(["31/02/2026", "OS-002", "Equipe A", "Manutenção", "Unidade A",
                         "Ana", "Concluído", 100, 80, 8, 0])
            wb_a.save(entrada / "equipe_a.xlsx")

            wb_b = Workbook()
            ws_b = wb_b.active
            ws_b.append(["Data execução", "OS", "Time", "Tipo de serviço", "Unidade",
                         "Funcionário", "Situação", "Meta prevista", "Produção", "Horas", "H. extras"])
            ws_b.append(["2026-09-02", "os-003", "equipe b", "inspecao", "unidade b",
                         "Beto", "finalizado", "1.200,50", "900,25", "7,5", "1,5"])
            ws_b.append(["03/09/2026", "OS-001", "Equipe B", "Inspeção", "Unidade B",
                         "Beto", "Pendente", 100, 0, 8, 0])
            wb_b.save(entrada / "equipe_b.xlsx")

            resultado = consolidar_pasta(entrada, saida)
            self.assertEqual((resultado.arquivos_lidos, resultado.linhas_lidas,
                              resultado.linhas_validas, resultado.linhas_descartadas,
                              resultado.inconsistencias), (2, 4, 2, 2, 2))

            dados = load_workbook(saida / "consolidado.xlsx", read_only=True, data_only=True)
            self.assertEqual(dados.active.max_row, 3)
            self.assertEqual(dados.active["D3"].value, "Inspeção")
            self.assertEqual(dados.active["H3"].value, 1200.5)
            self.assertEqual(dados.active["K3"].value, 1.5)
            dados.close()

            problemas = load_workbook(saida / "inconsistencias.xlsx", read_only=True, data_only=True)
            self.assertEqual({problemas.active["E2"].value, problemas.active["E3"].value},
                             {"Data", "Ordem de serviço"})
            problemas.close()

            resumo = load_workbook(saida / "resumo.xlsx", read_only=True, data_only=True)
            self.assertEqual(resumo.active["B10"].value, 980.25)
            self.assertEqual(resumo.sheetnames, ["Visão geral", "Por equipe", "Por serviço", "Por mês"])
            resumo.close()


if __name__ == "__main__":
    unittest.main()
