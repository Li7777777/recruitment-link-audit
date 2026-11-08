"""Offline documentation checks, not live CLI authentication or API tests."""
from pathlib import Path
import re
import shlex
import unittest

ROOT = Path(__file__).resolve().parents[1]
GUIDE = ROOT / "references" / "feishu-cli.md"


def commands():
    guide = GUIDE.read_text(encoding="utf-8")
    blocks = re.findall(r"^[ \t]*```bash[ \t]*\n(.*?)^[ \t]*```", guide, re.M | re.S)
    return [shlex.split(line.strip()) for block in blocks for line in block.splitlines()
            if line.strip() and not line.lstrip().startswith("#")]


class CliGuideTests(unittest.TestCase):
    def test_guide_is_bundled_and_referenced(self):
        self.assertTrue(GUIDE.is_file())
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("references/feishu-cli.md", skill)
        self.assertIn("lark-cli", skill)
        self.assertIn("--as user", skill)
        for path in [ROOT / "SKILL.md", *sorted((ROOT / "references").glob("*.md"))]:
            for link in re.findall(r"\]\(([^)]+)\)", path.read_text(encoding="utf-8")):
                if "://" not in link and not link.startswith("#"):
                    self.assertTrue((path.parent / link.split("#", 1)[0]).exists(), f"Broken local link: {path.name}: {link}")

    def test_data_commands_are_read_only_and_explicitly_user(self):
        allowed = {
            "base": {"+url-resolve", "+field-list", "+view-get", "+view-get-filter", "+record-list"},
            "sheets": {"+workbook-info"},
        }
        count = 0
        for args in commands():
            if args[0] != "lark-cli" or len(args) < 3 or args[1] not in allowed or args[2] == "--help":
                continue
            count += 1
            self.assertIn(args[2], allowed[args[1]])
            self.assertIn("--as", args)
            self.assertEqual(args[args.index("--as") + 1], "user")
        self.assertGreaterEqual(count, 6)

    def test_auth_examples_do_not_request_broad_permissions(self):
        found_scope_login = False
        for args in commands():
            self.assertNotIn("--recommend", args)
            self.assertNotIn("--domain", args)
            if args[:3] == ["lark-cli", "auth", "login"] and "--device-code" not in args:
                found_scope_login = True
                self.assertIn("--scope", args)
                self.assertIn("--no-wait", args)
                self.assertIn("--json", args)
        self.assertTrue(found_scope_login)

    def test_record_example_preserves_view_and_private_relative_artifact(self):
        examples = [args for args in commands() if args[:3] == ["lark-cli", "base", "+record-list"]]
        self.assertEqual(len(examples), 1)
        args = examples[0]
        for option in ["--view-id", "--offset", "--limit", "--output"]:
            self.assertIn(option, args)
        self.assertNotIn("--filter-json", args)
        self.assertNotIn("--page-all", args)
        self.assertEqual(args[args.index("--format") + 1], "ndjson")
        self.assertEqual(args[args.index("--offset") + 1], "0")
        self.assertFalse(Path(args[args.index("--output") + 1]).is_absolute())

    def test_pagination_scope_and_output_contract_caveats_remain_documented(self):
        guide = GUIDE.read_text(encoding="utf-8")
        for concept in ["next_offset", "has_more=false", "query_context", "view_filtered_records", "rev", "record_id", "ok == true", "--filter-json 会覆盖视图筛选", "--device-code"]:
            self.assertIn(concept, guide)
        self.assertIn("仓库和全局 Skill 之外", guide)
        self.assertIn("获得安装同意后", guide)


if __name__ == "__main__":
    unittest.main()
