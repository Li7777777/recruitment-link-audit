#!/usr/bin/env python3
"""Validate an evidence-backed recruitment audit and export exactly four Excel sheets.

This script does NOT browse, classify jobs, or infer employer application limits.
Usage: python export_excel.py audit.json -o audit.xlsx [--overwrite]
Dependency: openpyxl >= 3.1, < 4
"""
from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
from urllib.parse import urlsplit

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SHEETS = ["所有单位总览", "所有单位岗位总表", "其他信息表", "筛选后岗位表"]
KEPT_RESULTS = {"纳入", "待核实"}
RESULTS = {"纳入", "排除", "不纳入", "待核实"}
JOB_STATUSES = {"在招", "已停止", "暂未开放", "待核实"}
COMPANY_STATUSES = JOB_STATUSES | {"部分在招", "当前无岗位"}
COVERAGES = {"完整", "部分", "受阻", "未找到"}
ELIGIBILITIES = {"符合已知条件", "不符", "待核实", "未评估"}
QUOTA_STATUSES = {"已确认", "未核实", "冲突"}
CLASSIFICATION_BASES = {"岗位名称", "岗位职责", "招聘专业", "综合"}
CONFIDENCES = {"高", "中", "低"}
BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
COLORS = {"纳入": "E2F0D9", "排除": "FCE4D6", "不纳入": "E7E6E6", "待核实": "FFF2CC"}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def nonempty(obj, keys, where):
    for key in keys:
        require(isinstance(obj.get(key), str) and bool(obj[key].strip()), f"{where}.{key}: 必须为非空字符串")


def check_time(value, where):
    require(isinstance(value, str), f"{where}: 应为带时区的 ISO 8601 时间")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        raise ValueError(f"{where}: 无效时间 {value!r}") from None
    require(parsed.utcoffset() is not None, f"{where}: 时间必须包含时区，例如 +08:00")


def check_url(value, where, required=False):
    if not value and not required:
        return
    require(isinstance(value, str), f"{where}: 应为 URL 字符串")
    parsed = urlsplit(value)
    require(parsed.scheme in {"https", "http"} and bool(parsed.hostname)
            and not parsed.username and not parsed.password
            and not any(c.isspace() for c in value), f"{where}: 只接受不含账号密码的 http(s) URL")


def string_ids(values, where):
    require(isinstance(values, list) and all(isinstance(v, str) and v.strip() for v in values),
            f"{where}: 应为非空字符串组成的数组（数组可为空）")
    require(len(values) == len(set(values)), f"{where}: 存在重复 ID")


def validate_quota(quota, where):
    require(isinstance(quota, dict), f"{where}: 必须为对象，未知也必须记录")
    require(quota.get("status") in QUOTA_STATUSES, f"{where}.status: 无效值")
    maximum = quota.get("max_positions")
    require(maximum is None or (type(maximum) is int and maximum >= 0), f"{where}.max_positions: 应为非负整数或 null")
    require(type(quota.get("unlimited", False)) is bool, f"{where}.unlimited: 应为布尔值")
    require(not (maximum is not None and quota.get("unlimited")), f"{where}: 数值上限和无限制不能同时存在")
    if quota["status"] == "已确认":
        nonempty(quota, ["scope", "raw_text", "checked_at"], where)
        check_url(quota.get("source_url"), f"{where}.source_url", True)
        check_time(quota["checked_at"], f"{where}.checked_at")
        require(maximum is not None or quota.get("unlimited") or quota.get("shared_scope_id"),
                f"{where}: 确认的限投规则必须有数值上限、明确不限或共享/分层约束")
    else:
        require(maximum is None and not quota.get("unlimited"), f"{where}: 未核实/冲突的限投规则不能填确定数值")
    if quota.get("shared_scope_id"):
        nonempty(quota, ["shared_scope_note"], where)
    check_url(quota.get("source_url"), f"{where}.source_url")
    if quota.get("checked_at"):
        check_time(quota["checked_at"], f"{where}.checked_at")


def validate(data):
    require(isinstance(data, dict), "顶层必须为 JSON 对象")
    for key, typ in [("run", dict), ("companies", list), ("jobs", list), ("other_info", list)]:
        require(isinstance(data.get(key), typ), f"缺少或类型错误: {key}")
    run = data["run"]
    nonempty(run, ["scope", "scope_evidence", "rules_file", "rules_sha256", "checked_at"], "run")
    check_url(run.get("source_url"), "run.source_url", True)
    check_time(run["checked_at"], "run.checked_at")
    require(type(run.get("scope_verified")) is bool, "run.scope_verified: 必须为布尔值")
    require(re.fullmatch(r"[0-9a-fA-F]{64}", run["rules_sha256"]) is not None, "run.rules_sha256: 应为 SHA-256")
    for field in ["priority_order", "city_priority_order", "source_record_ids"]:
        string_ids(run.get(field), f"run.{field}")
    city_exclusions = run.get("city_exclusion_priorities", [])
    string_ids(city_exclusions, "run.city_exclusion_priorities")
    require(all(rank in run["city_priority_order"] for rank in city_exclusions),
            "城市排除等级必须属于city_priority_order")
    expected = run.get("source_expected_records")
    require(expected is None or (type(expected) is int and expected >= 0), "run.source_expected_records: 应为非负整数或 null")
    ignored = run.get("ignored_records", [])
    require(isinstance(ignored, list), "run.ignored_records: 应为数组")
    for item in ignored:
        require(isinstance(item, dict), "ignored_records 元素应为对象")
        nonempty(item, ["record_id", "reason"], "ignored_records")
    string_ids([i["record_id"] for i in ignored], "ignored_records.record_id")
    ignored_ids = {i["record_id"] for i in ignored}
    mapped = set(ignored_ids)
    companies = {}
    for company in data["companies"]:
        require(isinstance(company, dict), "companies 元素应为对象")
        nonempty(company, ["company_id", "name", "checked_at"], "company")
        cid = company["company_id"]
        require(cid not in companies, f"重复 company_id: {cid}")
        companies[cid] = company
        require(company.get("coverage") in COVERAGES, f"{cid}: 无效 coverage")
        require(company.get("status") in COMPANY_STATUSES, f"{cid}: 无效 status")
        check_time(company["checked_at"], f"{cid}.checked_at")
        string_ids(company.get("source_record_ids"), f"{cid}.source_record_ids")
        require(not ignored_ids.intersection(company["source_record_ids"]), f"{cid}: 已映射的源记录不能同时忽略")
        mapped.update(company["source_record_ids"])
        for field in ["source_url", "application_url", "status_source_url"]:
            check_url(company.get(field), f"{cid}.{field}")
        if company["status"] != "待核实":
            nonempty(company, ["status_evidence"], cid)
            check_url(company.get("status_source_url"), f"{cid}.status_source_url", True)
        if company["coverage"] != "完整":
            nonempty(company, ["notes"], cid)
        validate_quota(company.get("quota"), f"{cid}.quota")
    require(mapped == set(run["source_record_ids"]), "源记录对账失败：每条源记录必须映射到单位或附理由忽略；不得多出范围外记录")
    if "source_companies" in data:
        require(isinstance(data["source_companies"], list), "source_companies: 必须为数组")
        source_ids = set()
        for entry in data["source_companies"]:
            require(isinstance(entry, dict), "source_companies 元素应为对象")
            nonempty(entry, ["source_record_id", "name", "checked_at"], "source_company")
            rid = entry["source_record_id"]
            require(rid not in source_ids, f"重复原表公司记录: {rid}")
            source_ids.add(rid)
            require(entry.get("coverage") in COVERAGES, f"{rid}: 无效原表公司 coverage")
            check_time(entry["checked_at"], f"{rid}.checked_at")
            for field in ["source_url", "application_url"]:
                check_url(entry.get(field), f"{rid}.{field}")
            if entry["coverage"] != "完整":
                nonempty(entry, ["notes"], rid)
            if "quota" in entry:
                validate_quota(entry["quota"], f"{rid}.quota")
        require(source_ids == set(run["source_record_ids"]) - ignored_ids,
                "原表公司对账失败：source_companies必须逐条覆盖非忽略源记录，不可增加子公司或遗漏原表公司")
    row_ids = set()
    city_rows = set()
    grouped = defaultdict(list)
    for job in data["jobs"]:
        require(isinstance(job, dict), "jobs 元素应为对象")
        nonempty(job, ["row_id", "company_id", "job_key", "name", "city", "rationale", "checked_at"], "job")
        jid = job["row_id"]
        require(jid not in row_ids, f"重复 row_id: {jid}")
        row_ids.add(jid)
        require(job["company_id"] in companies, f"{jid}: 引用了不存在的单位")
        city_key = (job["company_id"], job["job_key"], job["city"])
        require(city_key not in city_rows, f"{jid}: 同一投递身份、城市重复，需先去重")
        city_rows.add(city_key)
        grouped[job["company_id"]].append(job)
        require(job.get("status") in JOB_STATUSES, f"{jid}: 无效岗位 status")
        require(job.get("result") in RESULTS, f"{jid}: 无效 result")
        require(job.get("eligibility") in ELIGIBILITIES, f"{jid}: 无效 eligibility")
        check_time(job["checked_at"], f"{jid}.checked_at")
        check_url(job.get("source_url"), f"{jid}.source_url", True)
        check_url(job.get("application_url"), f"{jid}.application_url")
        check_url(job.get("status_source_url"), f"{jid}.status_source_url", job["status"] != "待核实")
        check_url(job.get("deadline_source_url"), f"{jid}.deadline_source_url")
        if job["status"] != "待核实":
            nonempty(job, ["status_evidence"], jid)
        # Direction, final screening, city, and eligibility are separate dimensions.
        direction = job.get("direction_result", job["result"])
        require(direction in RESULTS, f"{jid}: 无效 direction_result")
        if direction == "纳入":
            require(job.get("priority") in run["priority_order"], f"{jid}: 方向纳入岗位缺少规则定义的岗位优先级")
        else:
            require(not job.get("priority"), f"{jid}: 方向未匹配时不能携带优先级")
        if job["result"] == "纳入":
            require(direction == "纳入", f"{jid}: 最终纳入必须有已匹配的方向")
        if job["result"] == "待核实" and direction == "纳入":
            nonempty(job, ["blocking_issues"], jid)
        # Empty city rank is accepted for legacy data, never required by exclusion.
        city_values = run["city_priority_order"] + ["待核实", ""] if run["city_priority_order"] else [""]
        require(job.get("city_priority", "") in city_values, f"{jid}: 城市优先级无效")
        if job.get("city_priority") in city_exclusions:
            require(job["result"] == "排除", f"{jid}: 命中城市排除规则，最终结果必须排除")
        if city_exclusions and job.get("city_priority", "") in {"", "待核实"}:
            require(job["result"] != "纳入", f"{jid}: 已启用地域排除，地点未核实不能最终纳入；可保留方向等级")
        for field in ["blocking_issues", "nonblocking_notes"]:
            require(isinstance(job.get(field, ""), str), f"{jid}.{field}: 应为字符串")
        if job.get("classification_basis"):
            require(job["classification_basis"] in CLASSIFICATION_BASES, f"{jid}: 无效 classification_basis")
        if job.get("classification_confidence"):
            require(job["classification_confidence"] in CONFIDENCES, f"{jid}: 无效 classification_confidence")
        if job["result"] == "排除":
            nonempty(job, ["blacklist_reason"], jid)
        if job["result"] == "待核实" or job["status"] == "待核实":
            require(bool(job.get("uncertainties") or job.get("blocking_issues")), f"{jid}: 待核实必须说明具体阻断事项")
        if job.get("city_priority") == "待核实":
            require(bool(job.get("uncertainties") or job.get("nonblocking_notes") or job.get("blocking_issues")),
                    f"{jid}: 地点待核实需有说明，但不改变方向结论")
    for cid, company in companies.items():
        require(company["source_record_ids"] or not grouped[cid], f"{cid}: 无原表来源的主体不能带岗位；请先补来源映射")
        active = any(j["status"] == "在招" for j in grouped[cid])
        if company["status"] in {"已停止", "当前无岗位", "暂未开放"}:
            require(not active, f"{cid}: 单位状态与仍在招岗位冲突")
        if company["status"] in {"在招", "部分在招"}:
            require(active, f"{cid}: 在招单位没有任何已确认在招的岗位记录")
    for info in data["other_info"]:
        require(isinstance(info, dict), "other_info 元素应为对象")
        nonempty(info, ["category", "item", "value"], "other_info")
        require(type(info.get("needs_review", False)) is bool, "other_info.needs_review: 必须为布尔值")
        require(not info.get("company_id") or info["company_id"] in companies, "other_info 引用了不存在的单位")
        if info.get("job_key"):
            require(any(j["job_key"] == info["job_key"] and
                        (not info.get("company_id") or j["company_id"] == info["company_id"])
                        for j in data["jobs"]), "other_info 引用了不存在的岗位")
        check_url(info.get("source_url"), "other_info.source_url")
        if info.get("checked_at"):
            check_time(info["checked_at"], "other_info.checked_at")
    return grouped


def unique_count(jobs, predicate=lambda j: True):
    return len({(j["company_id"], j["job_key"]) for j in jobs if predicate(j)})


def candidate(job):
    return job["status"] == "在招" and job["result"] == "纳入" and job["eligibility"] != "不符"


def flags(company, jobs):
    if not jobs:
        text = "无岗位记录"
    elif all(j["result"] == "排除" for j in jobs):
        text = "全部排除"
    elif all(j["result"] == "不纳入" for j in jobs):
        text = "全部不纳入"
    elif all(j["result"] in {"排除", "不纳入"} for j in jobs):
        text = "全部排除或不纳入"
    elif not any(candidate(j) for j in jobs):
        text = "无已确认在招的匹配候选"
    else:
        text = "有匹配候选"
    if company["coverage"] != "完整":
        return f"已见范围：{text}；核验未完成"
    return text


def official_max(quota):
    if quota["status"] != "已确认":
        return quota["status"]
    if quota.get("shared_scope_id"):
        cap = quota.get("max_positions")
        return f"单位≤{cap}，另受共享/分层限制" if cap is not None else "见共享/分层限投规则"
    return "明确不限岗位数量" if quota.get("unlimited") else quota["max_positions"]


def theoretical_max(company, jobs):
    count = unique_count(jobs, candidate)
    quota = company["quota"]
    if company["coverage"] != "完整":
        return f"未核实（已发现匹配在招 {count} 个）"
    if count == 0:
        return 0
    if company.get("has_internal_limits"):
        return f"需下属主体/分层额度统筹（匹配在招 {count} 个）"
    if quota.get("shared_scope_id"):
        return f"需共享/分层统筹（匹配在招 {count} 个）"
    if quota["status"] != "已确认":
        return f"未核实（匹配在招 {count} 个）"
    return count if quota.get("unlimited") else min(count, quota["max_positions"])


def source_entries(data):
    """One overview entry per actual source row; never invent a parent from child names."""
    if "source_companies" in data:
        return data["source_companies"]
    ignored = {row["record_id"] for row in data["run"].get("ignored_records", [])}
    actual = set(data["run"]["source_record_ids"]) - ignored
    require(not actual, "缺少原表公司名称映射，请补source_companies；即使只有一个核验主体也不能猜测原表名称")
    return []


def source_overview(entry, companies, grouped):
    """Aggregate audit subjects into the source row without adding their quota numbers."""
    rid = entry["source_record_id"]
    members = [c for c in companies if rid in c["source_record_ids"]]
    require(bool(members), f"{rid}: 原表公司没有对应核验主体")
    jobs = [job for c in members for job in grouped[c["company_id"]]]
    complete = entry["coverage"] == "完整" and all(c["coverage"] == "完整" for c in members)
    coverage = "完整" if complete else (
        "部分" if entry["coverage"] == "部分" or any(c["coverage"] in {"完整", "部分"} for c in members) else "受阻")
    states = {c["status"] for c in members}
    if any(j["status"] == "在招" for j in jobs):
        status = "在招" if states == {"在招"} and all(j["status"] == "在招" for j in jobs) else "部分在招"
    elif not complete or "待核实" in states:
        status = "待核实"
    elif states == {"已停止"}:
        status = "已停止"
    elif states == {"暂未开放"}:
        status = "暂未开放"
    else:
        status = "当前无岗位"
    same_unit = len(members) == 1 and members[0]["name"] == entry["name"]
    if "quota" in entry:
        quota = entry["quota"]
    elif same_unit:
        quota = members[0]["quota"]
    else:
        quota = {"status": "未核实", "max_positions": None, "unlimited": False, "scope": "原表公司整体招聘范围",
                 "notes": "整体限额未核实；下属主体额度不得相加或冒充原表公司的总限额"}
    single = members[0] if same_unit else {}
    internal_limits = not same_unit and any(c["quota"]["status"] == "已确认" and
        (c["quota"].get("max_positions") is not None or c["quota"].get("shared_scope_id")) for c in members)
    notes = entry.get("notes", "")
    if internal_limits:
        notes += "；下属主体有限额，理论上限需分层统筹，不能仅按整体限额套min公式。"
    if not complete:
        notes += "；内部招聘主体或入口尚未完整核验，汇总仅代表已见范围。"
    result = {"company_id": rid, "name": entry["name"], "source_record_ids": [rid], "coverage": coverage,
              "status": status, "quota": quota, "has_internal_limits": internal_limits,
              "checked_at": entry["checked_at"], "notes": notes.lstrip("；"),
              "summary": entry.get("summary", ""), "source_url": entry.get("source_url") or single.get("source_url", ""),
              "application_url": entry.get("application_url") or single.get("application_url", ""),
              "status_evidence": single.get("status_evidence") or "按原表公司范围内的岗位及招聘主体状态汇总；具体原文见岗位总表。",
              "status_source_url": single.get("status_source_url") or entry.get("application_url") or entry.get("source_url", "")}
    return result, jobs


def build_rows(data, grouped):
    run = data["run"]
    companies = {c["company_id"]: c for c in data["companies"]}
    entries = source_entries(data)
    source_names = {entry["source_record_id"]: entry["name"] for entry in entries}
    overview_headers = ["原表记录ID", "公司名称（原表）", "招聘状态", "核验完整性", "特殊标记", "投递总结", "已见独立岗位数", "岗位城市记录数",
                        "确认在招岗位数", "筛选匹配在招岗位数", "其中资格已确认数", "排除岗位数", "不纳入岗位数", "筛选待核实岗位数",
                        "官方最多可投岗位数", "限投适用范围", "共享额度标识", "共享/分层约束", "限投原文", "限投证据链接", "限投核验时间",
                        "按筛选的理论可投上限(未扣已投)", "理论上限说明", "官方投递入口", "招聘状态依据", "招聘状态证据链接",
                        "飞书来源链接", "来源记录ID", "核验时间", "备注"]
    overview = []
    overview_groups = [source_overview(entry, data["companies"], grouped) for entry in entries]
    for company, jobs in overview_groups:
        company["source_url"] = company.get("source_url") or run["source_url"]
        quota = company["quota"]
        count = unique_count(jobs, candidate)
        verified = unique_count(jobs, lambda j: candidate(j) and j["eligibility"] == "符合已知条件")
        summary = (f"{company['status']}；已见 {unique_count(jobs)} 个独立岗位；"
                   f"匹配且在招 {count} 个（资格已确认 {verified} 个）；{flags(company, jobs)}")
        if company.get("summary"):
            summary += "；" + company["summary"]
        notes = "未扣除个人历史已投；不是剩余志愿数或保证可投数。"
        if count != verified:
            notes += "包含资格未确认候选，需核实届次/学历/专业等。"
        if quota.get("shared_scope_id") or company.get("has_internal_limits"):
            notes += "共享/下层额度不可直接按主体相加，需分层统筹。"
        overview.append([company["company_id"], company["name"], company["status"], company["coverage"], flags(company, jobs), summary,
                         unique_count(jobs), len(jobs), unique_count(jobs, lambda j: j["status"] == "在招"), count, verified,
                         unique_count(jobs, lambda j: j["result"] == "排除"), unique_count(jobs, lambda j: j["result"] == "不纳入"),
                         unique_count(jobs, lambda j: j["result"] == "待核实"), official_max(quota), quota.get("scope", "待核实"),
                         quota.get("shared_scope_id", ""), quota.get("shared_scope_note", ""), quota.get("raw_text", ""),
                         quota.get("source_url", ""), quota.get("checked_at", ""), theoretical_max(company, jobs), notes,
                         company.get("application_url", ""), company.get("status_evidence", ""), company.get("status_source_url", ""),
                         company.get("source_url", ""), "; ".join(company["source_record_ids"]), company["checked_at"], company.get("notes", "")])
    job_fields = [("row_id", "城市记录ID"), ("company_id", "单位ID"), ("company_name", "单位名称"), ("job_key", "独立投递身份"),
                  ("job_id", "官方岗位ID"), ("name", "岗位名称"), ("department", "部门/子单位"), ("batch", "招聘批次"), ("city", "工作城市"),
                  ("status", "岗位状态"), ("result", "筛选结果"), ("priority", "岗位优先级"), ("city_priority", "城市优先级"),
                  ("rationale", "筛选原文依据"), ("blacklist_reason", "排除原因"), ("uncertainties", "待核实事项"),
                  ("duties", "岗位职责"), ("requirements", "任职要求"), ("majors", "招聘专业"), ("education", "学历要求"),
                  ("graduation", "届次要求"), ("eligibility", "个人资格核验"), ("headcount", "招聘人数(非限投数)"),
                  ("deadline", "截止时间"), ("deadline_basis", "截止时间依据/时区"), ("deadline_source_url", "截止证据链接"),
                  ("application_url", "投递链接"), ("source_url", "岗位来源链接"), ("status_evidence", "岗位状态原文"),
                  ("status_source_url", "岗位状态证据链接"), ("checked_at", "核验时间"), ("notes", "备注"),
                  ("direction_result", "方向判断"), ("classification_basis", "判断依据类型"),
                  ("classification_confidence", "判断置信度"), ("blocking_issues", "阻断事项"),
                  ("nonblocking_notes", "非阻断补充说明"), ("source_company_name", "原表公司名称"),
                  ("source_record_ids", "原表记录ID")]
    priorities = {p: i for i, p in enumerate(run["priority_order"])}
    cities = {p: i for i, p in enumerate(run["city_priority_order"])}
    fallback = len(priorities)
    ordered = sorted(data["jobs"], key=lambda j: (
        {"纳入": 0, "待核实": 1, "不纳入": 2, "排除": 3}[j["result"]],
        priorities.get(j.get("priority"), fallback), cities.get(j.get("city_priority"), len(cities)),
        companies[j["company_id"]]["name"], j["job_key"], j["city"]))
    job_rows = [[companies[j["company_id"]]["name"] if key == "company_name" else
                 "; ".join(source_names[rid] for rid in companies[j["company_id"]]["source_record_ids"]) if key == "source_company_name" else
                 "; ".join(companies[j["company_id"]]["source_record_ids"]) if key == "source_record_ids" else
                 j.get("direction_result", j["result"]) if key == "direction_result" else j.get(key, "")
                 for key, _ in job_fields] for j in ordered]
    info_headers = ["类别", "单位ID", "岗位投递身份", "事项", "内容", "来源链接", "原文/证据", "核验时间", "需复核"]
    info_rows = []

    def info(category, item, value, cid="", key="", url="", quote="", checked_at=None, review=False):
        info_rows.append([category, cid, key, item, str(value), url, quote, checked_at or run["checked_at"], "是" if review else "否"])

    complete = (run["scope_verified"] and all(c["coverage"] == "完整" for c in data["companies"])
                and all(c["coverage"] == "完整" for c, _ in overview_groups)
                and (run.get("source_expected_records") is None or run["source_expected_records"] == len(run["source_record_ids"])))
    info("运行元信息", "采集完整性", "完整采集（仍以待核实事项为准）" if complete else "部分完成：不得声称全量核验", review=not complete)
    for key, title in [("scope", "本次指定范围"), ("scope_verified", "视图/单位范围是否已核实"), ("scope_evidence", "范围核实依据"),
                       ("rules_file", "实际读取的规则文件"), ("rules_sha256", "规则文件 SHA-256"), ("priority_order", "岗位优先级顺序"),
                       ("city_priority_order", "独立城市优先级顺序"), ("city_exclusion_priorities", "硬性排除的城市等级"),
                       ("source_expected_records", "源视图预期记录数(非单位数)"),
                       ("notes", "运行说明")]:
        if key in run:
            value = json.dumps(run[key], ensure_ascii=False) if isinstance(run[key], list) else str(run[key])
            info("运行元信息", title, value, url=run["source_url"] if key == "scope" else "")
    info("统计口径", "总览公司行/内部核验主体/源记录/独立投递身份/城市记录数",
         f"{len(entries)} / {len(companies)} / {len(run['source_record_ids'])} / {unique_count(data['jobs'])} / {len(data['jobs'])}")
    info("统计口径", "总览与原表逐记录对齐", "公司名称和顺序以source_companies为准，一条非忽略源记录一行；子公司只在岗位明细和证据中展开，不合成“若干子公司”行。同名不同源记录不按名称合并；一个岗位关联多个重复源记录时，总览行之间不可直接相加。")
    info("统计口径", "独立投递身份", "相同 company_id + job_key 的多城市行只计一个岗位；官方独立岗位ID/志愿ID才可分开计数。各筛选结果计数可能因城市条件不同而重叠，不应相加。")
    info("统计口径", "最多可投 vs 招聘人数", "招聘人数不是限投岗位数；理论上限只对完整采集且已确认的独立单位限额计算，不扣历史已投，不保证个人资格；共享/分层额度不求和，不生成全局最大数。")
    info("统计口径", "筛选后岗位表", "从所有单位岗位总表仅按最终筛选结果保留纳入/待核实（待审核）行，移除不纳入/排除；列、内容、多城市粒度及相对顺序完全保留。不额外按招聘状态或个人资格删除，不代表保证当前可投。")
    info("统计口径", "招聘状态与筛选结论相互独立", "历史停招岗位可仍有方向优先级，但不计入匹配在招数；访问失败不等于已停招，方向匹配不等于个人资格满足。")
    for ignored in run.get("ignored_records", []):
        info("源记录对账", ignored["record_id"], ignored["reason"])
    for company in data["companies"]:
        cid = company["company_id"]
        quota = company["quota"]
        if not company["source_record_ids"]:
            info("未找到的指定单位", company["name"], company.get("notes") or "未定位到原表记录，仅记补充事项，不虚构总览公司行。", cid, review=True)
        elif company["coverage"] != "完整" or company["status"] == "待核实":
            info("异常与覆盖率", "单位核验未完成/状态待核实", company.get("notes") or "详见状态证据", cid, url=company.get("status_source_url", ""), review=True)
        if quota["status"] != "已确认" or quota.get("shared_scope_id") or quota.get("notes"):
            info("限投规则", quota.get("scope") or "适用范围待核实",
                 f"{quota['status']}；{quota.get('shared_scope_note', '')}；{quota.get('notes', '')}", cid,
                 url=quota.get("source_url", ""), quote=quota.get("raw_text", ""), checked_at=quota.get("checked_at"),
                 review=quota["status"] != "已确认" or bool(quota.get("shared_scope_id")))
    for company, _ in overview_groups:
        quota = company["quota"]
        if quota["status"] != "已确认" or quota.get("shared_scope_id"):
            info("原表公司整体限投", company["name"],
                 f"来源记录{company['company_id']}；{quota['status']}；{quota.get('shared_scope_note', '')}；{quota.get('notes', '')}",
                 url=quota.get("source_url", ""), quote=quota.get("raw_text", ""),
                 checked_at=quota.get("checked_at"), review=True)
    info("复核口径", "阻断事项与补充信息分开", "仅真实筛选阻断、明确冲突或候选开放状态不明等标需复核；个人资料未提供或名称直判尚无完整JD，不自动把岗位标待复核。城市独立评级；仅用户明确设为排除的城市等级才硬性排除，启用地域排除时地点未知会阻断最终纳入。")
    for company in data["companies"]:
        cid = company["company_id"]
        count = unique_count(grouped[cid], lambda j: candidate(j) and j["eligibility"] in {"待核实", "未评估"})
        if count:
            info("个人资格提示", "匹配候选的个人资格尚未评估", f"{count} 个独立匹配在招岗位尚需本人对照学历/届次等；这不改变已确定的方向，不逐岗位生成阻断复核。", cid)
    legacy_city_count = sum(bool(run["city_priority_order"]) and not j.get("city_priority") for j in ordered)
    if legacy_city_count:
        info("非阻断补充", "旧数据缺少独立城市评级", f"{legacy_city_count} 条岗位城市记录未填写city_priority；保留原始地点，不从岗位等级代填，也不静默套用新规则。新任务需逐行独立评级。")
    seen_issues = set()
    for job in ordered:
        blocking = []
        if job["result"] == "待核实":
            blocking.append(job.get("blocking_issues") or job.get("uncertainties") or "筛选所需信息未确认")
        elif job.get("blocking_issues"):
            blocking.append(job["blocking_issues"])
        if job["result"] in {"纳入", "待核实"} and job["status"] == "待核实":
            blocking.append("岗位是否仍开放待核实")
        if (run.get("city_exclusion_priorities") and job["result"] == "待核实"
                and job.get("city_priority", "") in {"", "待核实"}):
            blocking.append("实际工作地点未确认，无法排除地域黑名单")
        issue = "；".join(dict.fromkeys(blocking))
        signature = (job["company_id"], job["job_key"], issue)
        if issue and signature not in seen_issues:
            seen_issues.add(signature)
            info("岗位阻断复核", job["name"], issue, job["company_id"], job["job_key"],
                 job["source_url"], checked_at=job["checked_at"], review=True)
        # Original uncertainties stay visible in the job table, not an all-purpose review flag.
        note = job.get("nonblocking_notes", "")
        if job.get("city_priority") == "待核实" and not run.get("city_exclusion_priorities"):
            note = "实际工作地点待补充；已有岗位方向保留。" + note
        signature = (job["company_id"], job["job_key"], note)
        if note and signature not in seen_issues:
            seen_issues.add(signature)
            info("非阻断补充", job["name"], note, job["company_id"], job["job_key"],
                 job["source_url"], checked_at=job["checked_at"])

    for item in data["other_info"]:
        info(item["category"], item["item"], item["value"], item.get("company_id", ""), item.get("job_key", ""),
             item.get("source_url", ""), item.get("quote", ""), item.get("checked_at"), item.get("needs_review", False))
    job_headers = [title for _, title in job_fields]
    filtered_rows = [row for job, row in zip(ordered, job_rows) if job["result"] in KEPT_RESULTS]
    return [(overview_headers, overview), (job_headers, job_rows), (info_headers, info_rows),
            (job_headers, filtered_rows)]


def literal_cell(ws, row, column, value):
    require(isinstance(value, (str, int, float, bool)) or value is None, f"{ws.title} R{row}C{column}: 单元格不能为列表/对象")
    if isinstance(value, str):
        require(not BAD_XML.search(value), f"{ws.title} R{row}C{column}: 存在 XML 非法控制字符，请清理并保留原始证据文件")
        require(len(value) <= 32767, f"{ws.title} R{row}C{column}: 超过 Excel 单元格长度限制，请改为摘要+本地证据路径")
    cell = ws.cell(row, column, value)
    if isinstance(value, str):
        # Explicit string cells prevent formula injection, including =, +, -, @ prefixes.
        cell.data_type = "s"
        if value.startswith(("https://", "http://")) and not any(c.isspace() for c in value):
            check_url(value, f"{ws.title} R{row}C{column}")
            cell.hyperlink = value
            cell.font = Font(color="0563C1", underline="single", size=10)
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    return cell


def export_audit(data, output, overwrite=False):
    grouped = validate(data)
    tables = build_rows(data, grouped)
    output = Path(output).resolve()
    require(output.suffix.lower() == ".xlsx", "输出文件必须使用 .xlsx 扩展名")
    require(overwrite or not output.exists(), f"输出已存在，未覆盖: {output}（显式 --overwrite 才可覆盖）")
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, (headers, rows) in zip(SHEETS, tables):
        ws = workbook.create_sheet(name)
        is_job_sheet = name in {SHEETS[1], SHEETS[3]}
        for row_num, values in enumerate([headers] + rows, 1):
            for col_num, value in enumerate(values, 1):
                cell = literal_cell(ws, row_num, col_num, value)
                if row_num == 1:
                    cell.fill = PatternFill("solid", fgColor="203864")
                    cell.font = Font(color="FFFFFF", bold=True, size=10)
            if row_num > 1:
                color = None
                if name == SHEETS[0]:
                    color = "FCE4D6" if "全部" in str(values[4]) else ("FFF2CC" if values[3] != "完整" else None)
                elif is_job_sheet:
                    color = COLORS.get(values[10])
                    if values[9] != "在招" or values[21] == "不符":
                        color = "E7E6E6" if values[9] != "待核实" else "FFF2CC"
                elif values[-1] == "是":
                    color = "FFF2CC"
                if color:
                    for cell in ws[row_num]:
                        cell.fill = PatternFill("solid", fgColor=color)
        ws.freeze_panes = "F2" if is_job_sheet else "C2" if name == SHEETS[0] else "D2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.zoomScale = 85
        ws.row_dimensions[1].height = 42
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:1"
        for col, title in enumerate(headers, 1):
            width = 36 if any(k in title for k in ["链接", "入口", "原文", "依据", "总结", "约束", "事项", "职责", "要求", "说明", "备注", "内容"]) else 20
            ws.column_dimensions[get_column_letter(col)].width = width
    handle, temp_name = tempfile.mkstemp(prefix=".audit-", suffix=".xlsx", dir=output.parent)
    os.close(handle)
    try:
        workbook.save(temp_name)
        workbook.close()
        check = load_workbook(temp_name)
        try:
            require(check.sheetnames == SHEETS, "Excel 回读失败：Sheet 不符")
            for name, (headers, rows) in zip(SHEETS, tables):
                ws = check[name]
                require(ws.max_row == len(rows) + 1 and ws.max_column == len(headers), f"Excel 回读失败：{name} 维度不符")
                require(not any(cell.data_type == "f" for row in ws for cell in row), "Excel 出现意外公式")
        finally:
            check.close()
        if overwrite:
            os.replace(temp_name, output)
        else:
            # Exclusive reservation also prevents accidental overwrite if another run wins a race.
            with output.open("xb"):
                pass
            try:
                os.replace(temp_name, output)
            except OSError:
                output.unlink(missing_ok=True)
                raise
    finally:
        workbook.close()
        Path(temp_name).unlink(missing_ok=True)
    return {"output": str(output), "companies": len(tables[0][1]), "audit_subjects": len(data["companies"]), "job_city_rows": len(data["jobs"]),
            "unique_positions": len({(j["company_id"], j["job_key"]) for j in data["jobs"]}),
            "filtered_job_city_rows": len(tables[3][1]), "sheets": SHEETS}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="按 references/data-contract.md 生成的审计 JSON")
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    try:
        data = json.loads(args.input.read_text(encoding="utf-8-sig"))
        validate(data)
        # Check a locally accessible rule file against the recorded snapshot hash.
        rules = Path(data.get("run", {}).get("rules_file", ""))
        if rules.is_file():
            digest = hashlib.sha256(rules.read_bytes()).hexdigest()
            require(digest.lower() == data["run"].get("rules_sha256", "").lower(), "规则文件已变更：必须重新核对筛选或使用本次规则快照路径")
        print(json.dumps(export_audit(data, args.output, args.overwrite), ensure_ascii=False, indent=2))
    except (ValueError, OSError, KeyError, TypeError) as exc:
        parser.exit(2, f"导出失败: {exc}\n")


if __name__ == "__main__":
    main()
