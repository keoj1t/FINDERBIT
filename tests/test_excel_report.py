import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import main


class ExcelReportTests(unittest.TestCase):
    def test_builds_company_report_sheets_and_source_links(self) -> None:
        leads = [
            self.make_lead(
                project_name="Частная школа",
                city="Алматы",
                region="Алматы",
                object_type="Частная школа",
                work_type="new_construction",
                company_name="ТОО Пример",
                general_contractor="ТОО Строитель",
                stage="Начало строительства",
                timeline="2026-2027",
                offer_deadline="2026-11-24",
                budget="20 млрд ₸",
                project_scale="12 000 м²",
                clc_demand="Предположительно: панели и потолки",
                recommendation="Связаться с генподрядчиком",
                expected_result="Получить спецификацию отделки",
                risk_check="Подтвердить состав проекта",
                source_link="https://example.com/school",
                publication_date="2026-11-18",
                priority="A",
            ),
            self.make_lead(
                project_name="Офисный ремонт",
                city="Астана",
                region="Акмолинская область",
                object_type="Офис",
                work_type="reconstruction",
                company_name="ТОО Девелопер",
                general_contractor="Не установлено",
                stage="Проектирование",
                timeline="2027",
                clc_demand="Предположительно: напольные покрытия",
                recommendation="Найти подрядчика проекта",
                source_link="https://example.com/office",
                publication_date="2026-11-19",
                priority="B",
            ),
        ]
        report = main.normalize_report(
            {
                "leads": leads,
                "summary": {
                    "top_projects": ["Частная школа"],
                    "analytics": [
                        {
                            "project_name": "Частная школа",
                            "fact": "Школу строит ТОО Строитель.",
                            "impact": "Возможна поставка панелей.",
                            "action": "Связаться с генподрядчиком.",
                        }
                    ],
                    "priority_actions": ["Связаться с ТОО Строитель."],
                    "tech_news": "Новость о строительных технологиях.",
                    "technology_news": {
                        "date": "2026-11-19",
                        "initiative": "Новая технология",
                        "verified_fact": "Новость о строительных технологиях.",
                        "scale": "Масштаб проекта",
                        "clc_connection": "Потенциальная потребность в панелях",
                        "limitation": "Спецификация не опубликована",
                        "action": "Проверить требования проекта",
                        "source_link": "https://example.com/technology",
                    },
                },
            }
        )
        report["report_date"] = "2026-11-20"

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.xlsx"
            main.build_excel(report, path)
            workbook = load_workbook(path)

        self.assertEqual(
            workbook.sheetnames,
            [
                "Радар",
                "Проекты",
                "Образование и ремонты",
                "Отработать сегодня",
                "Технологии и ИИ",
            ],
        )
        self.assertEqual(
            workbook["Радар"].cell(row=4, column=4).value,
            "Сигнал",
        )
        self.assertEqual(
            workbook["Радар"].cell(row=5, column=4).value,
            "Частная школа",
        )
        self.assertEqual(
            workbook["Проекты"].cell(row=4, column=2).value,
            "Частная школа",
        )
        self.assertIn("12 000 м²", workbook["Проекты"].cell(row=4, column=7).value)
        self.assertEqual(
            workbook["Проекты"].cell(row=4, column=14).value,
            "2026-11-24",
        )
        self.assertEqual(
            workbook["Проекты"].cell(row=4, column=16).value,
            "Подтвердить состав проекта",
        )
        self.assertEqual(
            workbook["Образование и ремонты"].cell(row=4, column=2).value,
            "Частная школа",
        )
        self.assertEqual(
            workbook["Радар"].cell(row=5, column=9).hyperlink.target,
            "https://example.com/school",
        )
        self.assertEqual(
            workbook["Технологии и ИИ"].cell(row=4, column=3).value,
            "Новость о строительных технологиях.",
        )
        self.assertEqual(
            workbook["Технологии и ИИ"].cell(row=4, column=8).hyperlink.target,
            "https://example.com/technology",
        )
        self.assertEqual(
            workbook["Отработать сегодня"].cell(row=4, column=7).value,
            "Получить спецификацию отделки",
        )

        report_values = {
            cell.value
            for sheet in workbook.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if cell.value is not None
        }
        self.assertIn("CLC Company — деловая сводка, 20 ноября 2026 года", report_values)
        self.assertIn(
            "Связаться с ТОО Строитель.",
            workbook["Радар"].cell(row=2, column=1).value,
        )

    def test_saves_with_company_report_filename(self) -> None:
        report = {
            "report_date": "2026-11-20",
            "leads": [],
            "summary": {},
        }
        with tempfile.TemporaryDirectory() as directory:
            with (
                patch.object(main, "OUTPUT_DIR", Path(directory)),
                patch.object(main, "build_excel") as build_excel,
                patch.object(main.logger, "info"),
            ):
                path = main.save_outputs(report)

        self.assertEqual(path.name, "CLC_деловая_сводка_2026-11-20.xlsx")
        build_excel.assert_called_once_with(report, path)

    def test_builds_all_sheets_when_no_leads_are_found(self) -> None:
        report = main.normalize_report({"leads": [], "summary": {}})
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "empty.xlsx"
            main.build_excel(report, path)
            workbook = load_workbook(path)

        self.assertEqual(
            workbook.sheetnames,
            [
                "Радар",
                "Проекты",
                "Образование и ремонты",
                "Отработать сегодня",
                "Технологии и ИИ",
            ],
        )
        self.assertEqual(workbook["Радар"].cell(row=5, column=4).value, None)

    @staticmethod
    def make_lead(**overrides: str) -> dict[str, str]:
        lead = {key: main.NOT_SET for key in main.LEAD_FIELDS}
        lead.update(overrides)
        lead["crm_status"] = main.CRM_NOT_SENT
        return lead


if __name__ == "__main__":
    unittest.main()
