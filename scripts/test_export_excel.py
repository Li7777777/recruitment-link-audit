"""Offline tests: no browser, credentials, or real recruitment sites are used."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from openpyxl import load_workbook

from export_excel import SHEETS, export_audit, flags, theoretical_max, validate

ROOT = Path(__file__).resolve().parents[1]


def fixture():
    return json.loads((ROOT / "assets" / "example-audit.json").read_text(encoding="utf-8"))


def records(ws):
    rows = list(ws.values)
    return [dict(zip(rows[0], row)) for row in rows[1:]]


class ExportTests(unittest.TestCase):
    def test_three_sheets_and_city_deduplication(self):
        data = fixture()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "测试导出.xlsx"
            result = export_audit(data, path)
            self.assertEqual(result["companies"], 2)
            self.assertEqual(result["job_city_rows"], 3)
            self.assertEqual(result["unique_positions"], 2)
            wb = load_workbook(path)
            try:
                self.assertEqual(wb.sheetnames, SHEETS)
                a, b = records(wb[SHEETS[0]])
                self.assertEqual(a["已见独立岗位数"], 1)
                self.assertEqual(a["岗位城市记录数"], 2)
                self.assertEqual(a["筛选匹配在招岗位数"], 1)
                self.assertEqual(a["其中资格已确认数"], 0)
                self.assertEqual(a["官方最多可投岗位数"], 2)
                self.assertEqual(a["按筛选的理论可投上限(未扣已投)"], 1)
                self.assertIn("资格未确认", a["理论上限说明"])
                self.assertEqual(b["特殊标记"], "全部排除")
                self.assertEqual(b["官方最多可投岗位数"], "未核实")
                self.assertEqual(b["按筛选的理论可投上限(未扣已投)"], 0)
                self.assertEqual(len(records(wb[SHEETS[1]])), 3)
                for ws in wb:
                    self.assertTrue(ws.freeze_panes)
                    self.assertTrue(ws.auto_filter.ref)
                self.assertIsNotNone(wb[SHEETS[0]]["X2"].hyperlink)
            finally:
                wb.close()

    def test_formula_like_text_is_literal_and_no_formulas(self):
        data = fixture()
        malicious = '=HYPERLINK("https://example.invalid", "not a formula")'
        data["jobs"][0]["duties"] = malicious
        data["companies"][0]["name"] = "@untrusted"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "literal.xlsx"
            export_audit(data, path)
            wb = load_workbook(path)
            try:
                self.assertEqual(wb[SHEETS[1]]["Q2"].value, malicious)
                self.assertEqual(wb[SHEETS[1]]["Q2"].data_type, "s")
                self.assertFalse(any(c.data_type == "f" for ws in wb for row in ws for c in row))
            finally:
                wb.close()

    def test_shared_limits_not_independently_computed(self):
        data = fixture()
        company = data["companies"][0]
        company["quota"].update(shared_scope_id="fictional-group", shared_scope_note="集团合计≤3，每子公司≤2")
        jobs = validate(data)["example-a"]
        self.assertIn("共享/分层", theoretical_max(company, jobs))
        company["quota"]["max_positions"] = None
        validate(data)  # Group-only confirmed constraint is valid, not copied into a per-unit number.

    def test_unknown_quota_not_assumed_unlimited(self):
        data = fixture()
        company = data["companies"][0]
        company["quota"] = {"status": "未核实", "max_positions": None, "unlimited": False}
        self.assertIn("未核实", theoretical_max(company, validate(data)["example-a"]))
        company["quota"]["max_positions"] = 2
        with self.assertRaisesRegex(ValueError, "不能填确定数值"):
            validate(data)

    def test_closed_and_ineligible_jobs_do_not_count_as_candidates(self):
        data = fixture()
        data["companies"][0]["status"] = "已停止"
        for job in data["jobs"][:2]:
            job["status"] = "已停止"
        self.assertEqual(theoretical_max(data["companies"][0], validate(data)["example-a"]), 0)
        data = fixture()
        for job in data["jobs"][:2]:
            job["eligibility"] = "不符"
        self.assertEqual(theoretical_max(data["companies"][0], validate(data)["example-a"]), 0)

    def test_partial_all_excluded_only_claims_observed_scope(self):
        data = fixture()
        company = data["companies"][1]
        company.update(coverage="部分", notes="第二页读取失败")
        jobs = validate(data)["example-b"]
        self.assertIn("已见范围", flags(company, jobs))
        self.assertIn("核验未完成", flags(company, jobs))
        self.assertIn("未核实", theoretical_max(company, jobs))

    def test_empty_jobs_do_not_mean_all_excluded(self):
        data = fixture()
        company = data["companies"][1]
        company["status"] = "当前无岗位"
        data["jobs"] = data["jobs"][:2]
        jobs = validate(data)["example-b"]
        self.assertEqual(flags(company, jobs), "无岗位记录")
        self.assertNotIn("全部", flags(company, jobs))

    def test_mixed_excluded_and_not_included_marked(self):
        data = fixture()
        job = copy.deepcopy(data["jobs"][2])
        job.update(row_id="extra", job_key="different-identity", result="不纳入", blacklist_reason="")
        data["jobs"].append(job)
        self.assertEqual(flags(data["companies"][1], validate(data)["example-b"]), "全部排除或不纳入")
        for j in data["jobs"][2:]:
            j.update(result="不纳入", blacklist_reason="")
        self.assertEqual(flags(data["companies"][1], validate(data)["example-b"]), "全部不纳入")

    def test_source_record_accounting_and_group_expansion(self):
        data = fixture()
        data["companies"][1]["source_record_ids"].append("rec-fictional-a")
        validate(data)  # A group source record can produce several units.
        data["run"]["source_record_ids"].append("rec-lost")
        with self.assertRaisesRegex(ValueError, "源记录对账失败"):
            validate(data)
        data["run"]["ignored_records"].append({"record_id": "rec-lost", "reason": "说明行"})
        validate(data)
        data["run"]["ignored_records"].append({"record_id": "rec-fictional-a", "reason": "不能既映射又忽略"})
        with self.assertRaisesRegex(ValueError, "不能同时忽略"):
            validate(data)

    def test_missing_target_is_retained(self):
        data = fixture()
        missing = copy.deepcopy(data["companies"][1])
        missing.update(company_id="missing", name="指定但未找到的单位", source_record_ids=[],
                       coverage="未找到", status="待核实", notes="已检索指定视图，未找到该单位")
        data["companies"].append(missing)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "missing.xlsx"
            export_audit(data, path)
            wb = load_workbook(path)
            try:
                self.assertEqual(len(records(wb[SHEETS[0]])), 3)
                self.assertIn("部分完成", records(wb[SHEETS[2]])[0]["内容"])
            finally:
                wb.close()

    def test_duplicate_city_record_and_bad_references_rejected(self):
        data = fixture()
        data["jobs"][1]["city"] = "上海"
        with self.assertRaisesRegex(ValueError, "城市重复"):
            validate(data)
        data = fixture()
        data["jobs"][0]["company_id"] = "ghost"
        with self.assertRaisesRegex(ValueError, "不存在的单位"):
            validate(data)

    def test_no_city_ranking_and_custom_priority(self):
        data = fixture()
        data["run"]["city_priority_order"] = []
        data["run"]["priority_order"] = ["优先"]
        for job in data["jobs"][:2]:
            job.update(priority="优先", city_priority="")
        validate(data)

    def test_status_evidence_and_timezone_required(self):
        data = fixture()
        data["jobs"][0]["status_evidence"] = ""
        with self.assertRaisesRegex(ValueError, "status_evidence"):
            validate(data)
        data = fixture()
        data["jobs"][0]["checked_at"] = "2026-10-28T15:30:00"
        with self.assertRaisesRegex(ValueError, "时区"):
            validate(data)
        data = fixture()
        data["companies"][0]["status"] = "已停止"
        with self.assertRaisesRegex(ValueError, "仍在招岗位冲突"):
            validate(data)

    def test_no_overwrite_by_default(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "audit.xlsx"
            export_audit(fixture(), path)
            before = path.read_bytes()
            with self.assertRaisesRegex(ValueError, "输出已存在"):
                export_audit(fixture(), path)
            self.assertEqual(path.read_bytes(), before)
            export_audit(fixture(), path, overwrite=True)

    def test_overlong_text_and_xml_control_characters_not_silently_truncated(self):
        for value in ["x" * 32768, "bad\x00text"]:
            data = fixture()
            data["jobs"][0]["duties"] = value
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "invalid.xlsx"
                with self.assertRaises(ValueError):
                    export_audit(data, path)
                self.assertFalse(path.exists())

    def test_empty_audit_keeps_all_three_headers(self):
        data = fixture()
        data["companies"] = []
        data["jobs"] = []
        data["run"]["source_record_ids"] = []
        data["run"]["source_expected_records"] = 0
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "empty.xlsx"
            export_audit(data, path)
            wb = load_workbook(path)
            try:
                self.assertEqual(wb.sheetnames, SHEETS)
                self.assertEqual(wb[SHEETS[0]].max_row, 1)
                self.assertEqual(wb[SHEETS[1]].max_row, 1)
            finally:
                wb.close()

    def test_official_unlimited_and_explicit_zero(self):
        data = fixture()
        company = data["companies"][0]
        quota = company["quota"]
        quota.update(max_positions=None, unlimited=True, raw_text="虚构：本轮申请岗位数量不限")
        self.assertEqual(theoretical_max(company, validate(data)["example-a"]), 1)
        quota.update(max_positions=0, unlimited=False, raw_text="虚构：当前阶段本轮允许申请数量为0")
        self.assertEqual(theoretical_max(company, validate(data)["example-a"]), 0)

    def test_unverified_scope_and_count_mismatch_flag_partial(self):
        for field, value in [("scope_verified", False), ("source_expected_records", 99)]:
            data = fixture()
            data["run"][field] = value
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "partial.xlsx"
                export_audit(data, path)
                wb = load_workbook(path)
                try:
                    self.assertIn("部分完成", records(wb[SHEETS[2]])[0]["内容"])
                finally:
                    wb.close()

    def test_limit_caps_many_distinct_positions(self):
        data = fixture()
        for number in range(4):
            job = copy.deepcopy(data["jobs"][0])
            job.update(row_id=f"extra-{number}", job_key=f"distinct-{number}")
            data["jobs"].append(job)
        self.assertEqual(theoretical_max(data["companies"][0], validate(data)["example-a"]), 2)

    def test_user_priority_sorting_and_unknown_not_promoted(self):
        data = fixture()
        data["jobs"][0]["city_priority"] = "T1"
        lower = copy.deepcopy(data["jobs"][0])
        lower.update(row_id="lower-main", job_key="lower-main", priority="T1", city_priority="T0")
        pending = copy.deepcopy(data["jobs"][0])
        pending.update(row_id="pending", job_key="pending", result="待核实", priority="", city_priority="")
        data["jobs"].extend([lower, pending])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sort.xlsx"
            export_audit(data, path)
            wb = load_workbook(path)
            try:
                rows = records(wb[SHEETS[1]])
                self.assertEqual([r["主优先级"] for r in rows[:3]], ["T0", "T0", "T1"])
                self.assertEqual([r["城市子优先级"] for r in rows[:2]], ["T0", "T1"])
                self.assertEqual(rows[3]["筛选结果"], "待核实")
            finally:
                wb.close()
        pending["priority"] = "T0"
        with self.assertRaisesRegex(ValueError, "不能携带优先级"):
            validate(data)

    def test_cli_round_trip_and_rule_hash_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            data = fixture()
            data["run"]["rules_file"] = str(root / "规则.md")
            (root / "规则.md").write_text("test rule", encoding="utf-8")
            source = root / "audit.json"
            source.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            output = root / "output.xlsx"
            cmd = [sys.executable, str(ROOT / "scripts" / "export_excel.py"), str(source), "-o", str(output)]
            result = subprocess.run(cmd, capture_output=True)
            self.assertEqual(result.returncode, 2)
            self.assertFalse(output.exists())
            import hashlib
            data["run"]["rules_sha256"] = hashlib.sha256((root / "规则.md").read_bytes()).hexdigest()
            source.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            result = subprocess.run(cmd, capture_output=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue(output.exists())


if __name__ == "__main__":
    unittest.main()
