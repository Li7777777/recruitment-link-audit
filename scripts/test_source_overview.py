"""Source-row overview alignment, using fictional companies only."""
import copy
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from export_excel import SHEETS, build_rows, export_audit, validate

ROOT = Path(__file__).resolve().parents[1]


def fixture():
    return json.loads((ROOT / "assets" / "example-audit.json").read_text(encoding="utf-8"))


def group_fixture():
    data = fixture()
    data["run"].update(source_record_ids=["rec-fictional-group"], source_expected_records=1)
    data["source_companies"] = [{
        "source_record_id": "rec-fictional-group", "name": "示例集团原表名称（虚构）",
        "source_url": "https://example.invalid/feishu/group",
        "application_url": "https://example.invalid/group/campus",
        "coverage": "完整", "checked_at": data["run"]["checked_at"],
    }]
    for index, company in enumerate(data["companies"]):
        company.update(name=f"虚构下属主体{index + 1}", source_record_ids=["rec-fictional-group"])
    # Same local job ID in different subsidiaries must not collapse into one position.
    for job in data["jobs"]:
        job["job_key"] = "same-local-id"
    return data


def table_records(table):
    headers, rows = table
    return [dict(zip(headers, row)) for row in rows]


class SourceOverviewTests(unittest.TestCase):
    def test_group_expansion_keeps_one_source_row_and_all_job_details(self):
        data = group_fixture()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source-aligned.xlsx"
            result = export_audit(data, path)
            self.assertEqual(result["companies"], 1)
            self.assertEqual(result["audit_subjects"], 2)
            wb = load_workbook(path)
            try:
                overview = list(wb[SHEETS[0]].values)
                self.assertEqual(len(overview), 2)
                row = dict(zip(overview[0], overview[1]))
                self.assertEqual(row["原表记录ID"], "rec-fictional-group")
                self.assertEqual(row["公司名称（原表）"], data["source_companies"][0]["name"])
                self.assertEqual(row["已见独立岗位数"], 2)
                self.assertEqual(row["岗位城市记录数"], 3)
                self.assertEqual(row["筛选匹配在招岗位数"], 1)
                self.assertNotIn("若干子公司", row["投递总结"])
                jobs = list(wb[SHEETS[1]].values)
                detail = [dict(zip(jobs[0], values)) for values in jobs[1:]]
                self.assertEqual(len(detail), 3)
                self.assertEqual({j["单位名称"] for j in detail}, {"虚构下属主体1", "虚构下属主体2"})
                self.assertEqual({j["原表公司名称"] for j in detail}, {data["source_companies"][0]["name"]})
            finally:
                wb.close()

    def test_child_caps_are_not_summed_or_inherited_as_parent_cap(self):
        data = group_fixture()
        for company in data["companies"]:
            company["quota"] = copy.deepcopy(data["companies"][0]["quota"])
        row = table_records(build_rows(data, validate(data))[0])[0]
        self.assertEqual(row["官方最多可投岗位数"], "未核实")
        self.assertNotIsInstance(row["按筛选的理论可投上限(未扣已投)"], int)

    def test_parent_cap_requires_its_own_evidence_and_respects_nested_limits(self):
        data = group_fixture()
        quota = copy.deepcopy(data["companies"][0]["quota"])
        quota.update(max_positions=3, scope="虚构集团本轮整体上限", raw_text="虚构：集团本轮累计最多3个岗位")
        data["source_companies"][0]["quota"] = quota
        row = table_records(build_rows(data, validate(data))[0])[0]
        self.assertEqual(row["官方最多可投岗位数"], 3)
        self.assertIn("分层", row["按筛选的理论可投上限(未扣已投)"])
        # An overall-only quota can use min only when no lower-level limit is known.
        for company in data["companies"]:
            company["quota"] = {"status": "未核实", "max_positions": None, "unlimited": False}
        row = table_records(build_rows(data, validate(data))[0])[0]
        self.assertEqual(row["按筛选的理论可投上限(未扣已投)"], 1)
        del data["source_companies"][0]["quota"]["raw_text"]
        with self.assertRaisesRegex(ValueError, "raw_text"):
            validate(data)

    def test_source_group_partial_even_if_all_known_children_are_complete(self):
        data = group_fixture()
        data["source_companies"][0].update(coverage="部分", notes="仍有入口未完成核验")
        for job in data["jobs"]:
            job.update(result="排除", priority="", blacklist_reason="虚构排除条件")
        tables = build_rows(data, validate(data))
        row = table_records(tables[0])[0]
        self.assertEqual(row["核验完整性"], "部分")
        self.assertIn("已见范围", row["特殊标记"])
        self.assertIn("全部排除", row["特殊标记"])
        self.assertIn("部分完成", table_records(tables[2])[0]["内容"])

    def test_same_named_source_rows_remain_distinct_and_in_source_order(self):
        data = fixture()
        for entry in data["source_companies"]:
            entry["name"] = "同名原表公司（虚构）"
        data["source_companies"].reverse()
        rows = table_records(build_rows(data, validate(data))[0])
        self.assertEqual(len(rows), 2)
        self.assertEqual([row["原表记录ID"] for row in rows], ["rec-fictional-b", "rec-fictional-a"])

    def test_source_mapping_must_cover_exactly_actual_nonignored_records(self):
        for mode in ["missing", "duplicate", "invented"]:
            data = fixture()
            if mode == "missing":
                data["source_companies"].pop()
            elif mode == "duplicate":
                data["source_companies"].append(copy.deepcopy(data["source_companies"][0]))
            else:
                data["source_companies"][0]["source_record_id"] = "invented-child-row"
            with self.assertRaises(ValueError):
                validate(data)

    def test_legacy_ambiguous_subsidiaries_cannot_become_overview_rows(self):
        data = group_fixture()
        data.pop("source_companies")
        grouped = validate(data)
        with self.assertRaisesRegex(ValueError, "请补source_companies"):
            build_rows(data, grouped)
        data = fixture()
        data.pop("source_companies")
        with self.assertRaisesRegex(ValueError, "请补source_companies"):
            build_rows(data, validate(data))  # Even a sole audit subject might be a renamed subsidiary.

    def test_empty_group_is_not_marked_all_excluded(self):
        data = group_fixture()
        data["jobs"] = []
        for company in data["companies"]:
            company["status"] = "当前无岗位"
        row = table_records(build_rows(data, validate(data))[0])[0]
        self.assertEqual(row["招聘状态"], "当前无岗位")
        self.assertEqual(row["已见独立岗位数"], 0)
        self.assertEqual(row["特殊标记"], "无岗位记录")


if __name__ == "__main__":
    unittest.main()
