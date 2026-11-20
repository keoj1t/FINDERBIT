import argparse
import unittest
from unittest.mock import patch

import main


class SingleSearchRunTests(unittest.TestCase):
    @patch.object(main.logger, "info")
    @patch.object(main, "notify_telegram")
    @patch.object(main, "save_run_report")
    @patch.object(
        main, "process_report", return_value=({"leads": [], "summary": {}}, None)
    )
    @patch.object(main, "get_raw_report", return_value={"leads": [], "summary": {}})
    @patch.object(main, "get_source_text", return_value="search results")
    def test_run_searches_and_extracts_once(
        self,
        get_source_text,
        get_raw_report,
        _process_report,
        _save_report,
        _notify,
        _log_info,
    ) -> None:
        args = argparse.Namespace(
            from_json=None,
            input=None,
            skip_bitrix=True,
            dry_run=False,
            duplicates="update",
        )

        stats = main.run_once(args)

        self.assertEqual(stats.errors, 0)
        get_source_text.assert_called_once_with(stats)
        get_raw_report.assert_called_once_with(
            "search results", main.load_text(main.PROMPT_FILE), stats
        )

    @patch.object(main, "save_search_text")
    @patch.object(main, "render_search_prompt", return_value="one search prompt")
    @patch.object(main, "search_projects_text", return_value="search results")
    def test_source_text_uses_one_search(
        self, search_projects_text, render_search_prompt, save_search_text
    ) -> None:
        stats = main.RunStats()

        text = main.get_source_text(stats)

        self.assertEqual(text, "search results")
        self.assertEqual(stats.errors, 0)
        render_search_prompt.assert_called_once_with()
        search_projects_text.assert_called_once_with("one search prompt")
        save_search_text.assert_called_once_with("search results")


if __name__ == "__main__":
    unittest.main()
