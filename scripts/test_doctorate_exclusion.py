"""Opt-in doctoral minimum-degree blacklist; all fixtures are fictional."""
import json
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from export_excel import SHEETS, build_rows, export_audit, validate

ROOT = Path(__file__).resolve().parents[1]


def fixture():
    return json.loads((ROOT / "assets" / "example-audit.json").read_text(encoding="utf-8"))


class DoctorateExclusionTests(unittest.TestCase):
    def test_required_doctorate_excludes_even_t0_and_is_absent_from_filtered_sheet(self):
        data = fixture()
        for job in data["jobs"][:2]:
            job.update(education="博士研究生", doctorate_requirement="必需", doctorate_evidence="虚构：最低学历博士研究生")
        with self.assertRaisesRegex(ValueError, "明确博士必需"):
            validate(data)
        for job in data["jobs"][:2]:
            job.update(result="排除", direction_result="纳入", blacklist_reason="最低学历明确为博士，命中博士必需黑名单")
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "doctoral-exclusion.xlsx"
            result = export_audit(data, output)
            self.assertEqual(result["filtered_job_city_rows"], 0)
            wb = load_workbook(output)
            try:
                overview = list(wb[SHEETS[0]].values)
                source_row = dict(zip(overview[0], overview[1]))
                self.assertEqual(source_row["特殊标记"], "全部排除")
                self.assertEqual(source_row["筛选匹配在招岗位数"], 0)
                jobs = list(wb[SHEETS[1]].values)
                self.assertEqual(len(jobs), 4)  # All original job-city rows remain in the full table.
                row = dict(zip(jobs[0], jobs[1]))
                self.assertEqual(row["岗位优先级"], "T0")
                self.assertEqual(row["博士是否必需"], "必需")
                self.assertEqual(row["筛选结果"], "排除")
                self.assertEqual(wb[SHEETS[3]].max_row, 1)
            finally:
                wb.close()

    def test_preferred_or_non_doctoral_options_are_not_blocked(self):
        for text in ["博士优先", "PhD preferred", "硕士及以上", "本科及以上", "硕士或博士", "本硕博均可"]:
            with self.subTest(text=text):
                data = fixture()
                for job in data["jobs"][:2]:
                    job.update(education=text, doctorate_requirement="非必需", doctorate_evidence=text)
                tables = build_rows(data, validate(data))
                self.assertEqual(len(tables[3][1]), 2)
                self.assertTrue(all(row[10] == "纳入" for row in tables[3][1]))

    def test_unprovided_degree_does_not_create_blanket_pending_reviews(self):
        data = fixture()
        for job in data["jobs"][:2]:
            for field in ["education", "doctorate_requirement", "doctorate_evidence"]:
                job.pop(field, None)
        tables = build_rows(data, validate(data))
        self.assertEqual(len(tables[3][1]), 2)
        self.assertFalse(any(row[0] == "岗位阻断复核" for row in tables[2][1]))
        column = tables[1][0].index("博士是否必需")
        self.assertEqual(tables[1][1][0][column], "未评估")

    def test_real_minimum_degree_conflict_blocks_final_acceptance_but_keeps_direction(self):
        data = fixture()
        data["jobs"][0].update(doctorate_requirement="待核实", doctorate_evidence="虚构：学历字段博士，岗位说明硕士，适用范围冲突")
        with self.assertRaisesRegex(ValueError, "博士门槛存在真实歧义"):
            validate(data)
        data["jobs"][0].update(result="待核实", direction_result="纳入", blocking_issues="博士最低门槛真实冲突待确认")
        tables = build_rows(data, validate(data))
        self.assertEqual(data["jobs"][0]["priority"], "T0")
        self.assertTrue(any(row[0] == "岗位阻断复核" and row[-1] == "是" for row in tables[2][1]))
        self.assertEqual(len(tables[3][1]), 2)  # Pending rows remain visible for review.

    def test_exclusion_is_opt_in_and_flag_must_be_boolean(self):
        data = fixture()
        data["run"]["exclude_doctorate_required"] = False
        for job in data["jobs"][:2]:
            job.update(doctorate_requirement="必需", doctorate_evidence="虚构：博士必需", education="博士")
        self.assertEqual(len(build_rows(data, validate(data))[3][1]), 2)
        data["run"]["exclude_doctorate_required"] = "false"
        with self.assertRaisesRegex(ValueError, "必须为布尔值"):
            validate(data)

    def test_confirmed_or_conflicting_doctorate_tag_requires_evidence(self):
        for tag in ["必需", "待核实"]:
            data = fixture()
            data["jobs"][0].update(doctorate_requirement=tag, doctorate_evidence="")
            with self.assertRaisesRegex(ValueError, "doctorate_evidence"):
                validate(data)

    def test_published_policy_defines_required_versus_preferred_boundary(self):
        rules = (ROOT / "岗位优先级.md").read_text(encoding="utf-8")
        section = rules.split("### 4. 明确要求博士学历的岗位", 1)[1].split("## 五、", 1)[0]
        for text in ["直接排除", "博士及以上", "PhD required", "博士优先", "PhD preferred", "硕士及以上", "本硕博均可"]:
            self.assertIn(text, section)
        self.assertIs(fixture()["run"]["exclude_doctorate_required"], True)


if __name__ == "__main__":
    unittest.main()
