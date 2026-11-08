"""Check the bundled, explicitly opt-in policy; no live data or semantic API calls."""
import json
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]


class DecisionPolicyTests(unittest.TestCase):
    def test_independent_city_bands_match_the_published_policy(self):
        rules = (ROOT / "岗位优先级.md").read_text(encoding="utf-8")
        bands = dict(re.findall(r"^\| (T[0-3])(?:（排除）)? \| (.+?) \|$", rules, re.M))
        self.assertEqual(set(bands), {"T0", "T1", "T2"})
        for region in ["江苏", "浙江", "上海"]:
            self.assertIn(region, bands["T0"])
        for region in ["新疆", "西藏", "内蒙古", "山西", "北京", "青海"]:
            self.assertIn(region, bands["T2"])
        self.assertIn("最终筛选排除", bands["T2"])
        fixture = json.loads((ROOT / "assets" / "example-audit.json").read_text(encoding="utf-8"))
        self.assertEqual(fixture["run"]["city_priority_order"], ["T0", "T1", "T2"])
        self.assertEqual(fixture["run"]["city_exclusion_priorities"], ["T2"])
        self.assertEqual(fixture["run"]["priority_order"], ["T0", "T1", "T2", "T3"])
        self.assertIn("### T3：明确接受计算机类专业", rules)
        self.assertIn("西安不是山西", rules)
        self.assertIn("方向T0 / 城市T2", rules)
        self.assertIn("最终result=排除", rules)

    def test_name_evidence_and_nonblocking_policy_are_explicit(self):
        skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        policy = (ROOT / "references" / "decision-policy.md").read_text(encoding="utf-8")
        self.assertIn("references/decision-policy.md", skill)
        for item in ["Java开发", "NLP算法", "客户经理", "嵌入式", "专业", "blocking_issues", "nonblocking_notes"]:
            self.assertIn(item, policy)
        self.assertIn("不自动按名称重写历史报告", policy)
        self.assertIn("不能仅凭岗位名称猜测是否在招", (ROOT / "岗位优先级.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
