# 审计 JSON 与导出契约

`scripts/export_excel.py` 只消费已经核验和判定的数据。它不能替代浏览器、判断网页真实性、自动读取飞书或把岗位按语义分级。导出校验通过也不证明事实已核验；事实完整性由执行此 Skill 的 Agent 负责。

## 1. 文件与运行

建议每轮保存：

```text
导出/招聘核验_YYYYMMDD_HHMMSS/
├── audit.json                 # UTF-8，持续保存结果
├── 规则快照.md                 # 用户本次规则原始字节
├── 来源台账.json               # 视图/源字段/原始记录 ID，可带非敏感快照
├── evidence/                  # 必要网页原文、截图或附件
└── 招聘核验_YYYYMMDD_HHMMSS.xlsx
```

先读取规则并按原始字节计算 SHA-256，例如：

```bash
python -c "import hashlib,pathlib; p=pathlib.Path(r'<用户规则绝对路径>'); print(hashlib.sha256(p.read_bytes()).hexdigest())"
```

新任务的JSON顶层必须有`run`（对象）、`source_companies`（原表公司统计项数组）、`companies`（内部核验主体数组）、`jobs`（数组）、`other_info`（数组）。时间统一 ISO 8601 **含时区**，如 `2026-10-28T15:30:00+08:00`（只是格式示例，运行时使用实际时间）。链接只支持无内嵌账号密码的 http/https；登录态 token 不应进入 JSON/报告。

所有文本使用字符串，未知数字用 null、未知文本用空字符串或明确“待核实”，不能用 0 冒充未知。数组/嵌套对象不得直接当 Excel 单元格；需要额外信息时写入 other_info 的文本字段。未在下列字段中定义的额外 JSON 字段不会自动导出，重要信息应放 notes 或 other_info。

## 2. run：本次范围和规则

| 字段 | 类型/要求 |
| --- | --- |
| source_url | 必填，飞书表/视图 URL |
| scope | 必填，本次用户指定视图/单位/交集的文字范围 |
| scope_verified | 必填布尔值，是否已验证实际取得记录确属指定范围 |
| scope_evidence | 必填，视图筛选核实、分页结束、UI/API ID 对照等依据；失败也需写原因 |
| source_record_ids | 必填字符串数组，实际读到的范围内源行 ID，去重；普通表可用稳定行键 |
| source_expected_records | 非负整数或 null，来源显示的预期**源记录数**，不是单位数，不得拿整表数量充当视图数量 |
| ignored_records | 可选数组，每项 {record_id, reason}，仅用于已读到但不属于招聘单位的空行/说明行 |
| rules_file | 必填，本次实际使用的规则绝对路径（推荐指向原始字节快照），而非 Skill 自带默认规则 |
| rules_sha256 | 必填，64 位十六进制 SHA-256 |
| checked_at | 必填，本次范围核验时间，含时区 |
| priority_order | 必填字符串数组，从高到低，按本次规则，如 ["T0","T1","T2","T3"]；不是脚本内置等级 |
| city_priority_order | 必填字符串数组，独立城市等级从高到低；随附规则为 ["T0","T1","T2"]，不设置城市T3；其他规则按其定义，未设城市等级时为 [] |
| city_exclusion_priorities | 可选字符串数组，默认[]，必须属于city_priority_order；随附规则为["T2"]，只有用户明确设定的等级才作为硬性地域排除 |
| notes | 可选，本轮限制/快照路径等 |

`source_record_ids` 必须恰好等于所有 company.source_record_ids 与 ignored_records.record_id 的并集。可一源行展开多个单位，但不能把同一行同时映射和忽略，不能混入范围外行。只存在于用户单位名单而源表未找到的单位可用空来源 ID。

若 scope_verified=false、任一单位 coverage 不是完整，或源预期数与实际 ID 数不一致，脚本会把本次报告标“部分完成”。没有预期数量时需要用其他结束证据核实完整性。

脚本命令行在 rules_file 能本地访问时会复核哈希，不匹配则拒绝导出。跨机器路径不可访问时无法自动复核，Agent 仍须确保规则快照真实。

## 2.1 source_companies：与原表对齐的总览公司统计项

**新任务必须提供。** 每条非忽略原表记录恰好一项，按原视图顺序排列，source_record_id唯一且集合等于run.source_record_ids减去ignored_records。总览只由此数组生成，不直接遍历内部companies输出行。

| 字段 | 要求 |
| --- | --- |
| source_record_id | 必填，实际原表记录ID/稳定行键，不是人工生成的子公司ID |
| name | 必填，原表公司名称原文；不改成子公司名或“若干子公司”占位，不按名称合并不同源记录 |
| coverage | 必填，完整/部分/受阻/未找到；描述整个原表公司范围，不能只代表已知子公司都完成 |
| checked_at | 必填，本次原表公司范围真实核验时间，含时区 |
| source_url | 建议提供，定位原表记录的URL；缺失时可回退本次视图链接 |
| application_url | 可选，原表公司整体投递入口，不用某个子公司入口冒充 |
| quota | 可选，与下述quota契约相同，但适用范围是该原表公司整体；已确认需独立原文与证据 |
| summary | 可选，对整个原表公司作投递总结，不新增子公司统计行 |
| notes | coverage非完整时必填，说明范围未完成/失败原因 |

companies.source_record_ids决定各招聘主体归属于哪些原表统计项。一条原表记录可展开多个子公司，它们的岗位汇总到同一行，总览仍只有原表公司原名。原表本身独立列出的子公司仍作为其自己的原表行保留。同名不同记录不自动合并；同一主体若对应重复源记录，每条原表行仍保留，但行间统计不能直接累加。

整体quota未提供时，只有一个核验主体且名称与原表完全一致的简单情况才复用该主体额度；其他情况标整体限额未核实，不取某个子公司数值、不将子公司数值相加。有下属主体数值/共享额度约束时，即使整体上限已确认，理论可投上限也需分层统筹，不直接套整体min公式。

虚构示例：

```json
"source_companies": [{
  "source_record_id": "rec-fictional-group",
  "name": "示例集团原表名称（虚构）",
  "source_url": "https://example.invalid/feishu/group",
  "coverage": "完整",
  "checked_at": "2026-10-28T15:30:00+08:00"
}]
```

旧JSON未提供此数组且存在非忽略源记录时拒绝导出，即使只有一个核验主体也不能保证它的名字就是原表公司名。需从来源台账/原表补回实际名称映射，不根据门户或子公司名猜集团名。未定位到任何原表记录的额外指定单位放其他信息表，不虚构总览行；无来源主体也不能携带岗位。

## 3. companies：内部核验主体（不等于总览行）

每个实际招聘主体/轮次/额度单元一项，可以是原表公司或其子公司。不得只保存有推荐岗位的主体；但不把每个主体直接显示为总览公司行。

| 字段 | 类型/要求 |
| --- | --- |
| company_id | 必填，稳定字符串 ID；全文件唯一 |
| name | 必填单位名称，必要时附轮次以区分同名条目 |
| source_record_ids | 必填字符串数组，合并后对应的所有源行 ID，可为空 |
| source_url | 可选，定位飞书来源的 URL |
| application_url | 可选，官方投递入口，未找到可为空 |
| coverage | 必填：完整 / 部分 / 受阻 / 未找到 |
| status | 必填：在招 / 部分在招 / 已停止 / 暂未开放 / 当前无岗位 / 待核实 |
| status_evidence | 除待核实外必填，针对正确主体/本轮的关键原文或确证的列表情况说明 |
| status_source_url | 除待核实外必填，状态证据 URL |
| checked_at | 必填，核验该单位的真实时间 |
| quota | 必填对象，未知也必须给 status=未核实，见下表 |
| summary | 可选，具体投递建议、优先岗位、操作提醒；脚本另自动生成客观计数总结 |
| notes | 可选；coverage 非完整时必填，写未完成/未找到原因与范围 |

status 为在招/部分在招，至少有一条该单位的 job.status=在招；已停止/暂未开放/当前无岗位不能同时存在已确认在招岗位。未知详情不可伪造为停招。coverage=完整仅表示本次范围采集核验的完整性，不表示所有个人资格和限额都已确认。

### quota：官方限投规则

| 字段 | 类型/要求 |
| --- | --- |
| status | 必填：已确认 / 未核实 / 冲突 |
| max_positions | 非负整数或 null；**本单位本轮可独立应用的上限**，不是集团总上限、招聘人数或个人剩余数 |
| unlimited | 布尔值，默认 false；仅官方明确不限岗位数量时 true，与数值上限互斥 |
| scope | 已确认时必填，原文规则适用的招聘主体/批次/人群/时段 |
| raw_text | 已确认时必填，限投原文；冲突建议合并双方原文并在 other_info 分开存证 |
| source_url | 已确认时必填，权威规则 URL |
| checked_at | 已确认时必填，带时区核验时间；其他状态如提供也须有效 |
| shared_scope_id | 可选，共享/互斥/复杂分层约束的稳定标识。存在时脚本不套独立额度 min 公式 |
| shared_scope_note | shared_scope_id 存在时必填，明确各层上限及共享关系，不单靠编号 |
| notes | 可选，未知查找过程、冲突、优先志愿/撤回规则等 |

已确认时必须至少有：独立数值上限、明确不限，或 shared_scope_id 描述复杂约束。未核实/冲突时 max_positions 必须为 null 且 unlimited=false。只证实集团共享≤3时，max_positions=null，shared_scope_note 写“集团本轮合计最多3，每单位独立限额尚不明确”等真实含义。

未知示例：

```json
{"status":"未核实","max_positions":null,"unlimited":false,"scope":"本轮单位额度待核实","notes":"已查公告与FAQ，未找到限投说明"}
```

分层示例（虚构，需有真实证据才能用）：

```json
{
  "status":"已确认",
  "max_positions":2,
  "unlimited":false,
  "scope":"某子公司本轮校招；同时受集团额度约束",
  "raw_text":"每家子公司最多2个岗位，集团内累计最多3个岗位。",
  "source_url":"https://example.invalid/official-faq",
  "checked_at":"2026-10-28T15:30:00+08:00",
  "shared_scope_id":"某集团-本轮",
  "shared_scope_note":"每子公司≤2，集团合计≤3；不可将各单位额度相加"
}
```

## 4. jobs：本范围全部已见岗位-城市行

必填字段：

- `row_id`：唯一的城市记录 ID。
- `company_id`：关联单位，必须在 companies 中。
- `job_key`：该核验主体下的真实独立投递身份。同一岗位多个可选城市共用此值；官网明确独立志愿ID才分开。跨子公司汇总时按(company_id, job_key)去重，不能仅凭相同本地岗位ID合并。
- `name`：原始岗位名称。
- `city`：实际城市；未知写“待核实”，多城市按城市拆行。相同 company_id+job_key+city 不得重复。
- `status`：在招 / 已停止 / 暂未开放 / 待核实。
- `result`：纳入 / 排除 / 不纳入 / 待核实。
- `rationale`：本次规则对应的岗位名称/职责/专业等原文与判定理由。明确名称可以直判，不强制完整JD；待核实需说明真正会改变结论的证据缺口。
- `eligibility`：符合已知条件 / 不符 / 待核实 / 未评估；未知不能写符合。
- `source_url`：岗位详情/本轮岗位公告等证据 URL。
- `checked_at`：实际岗位核验时间。

有条件必填：

- status 非待核实：`status_evidence` + `status_source_url`，说明当前是否开放及依据。
- `direction_result` 可选，取纳入/排除/不纳入/待核实；未提供时兼容旧数据，等于result。它描述方向，result描述最终筛选。
- direction_result=纳入：`priority` 必须存在于run.priority_order；方向未匹配/未知时priority留空。最终result=纳入必须有匹配方向。
- 最终result=待核实但direction_result=纳入时，保留priority，必须填写`blocking_issues`解释真实门槛，如正式校招类型未明。这类岗位仍不计候选，并排在最终纳入之后。
- result=排除：`blacklist_reason`必填，说明黑名单或招聘类型等范围排除原因。方向等级通常留空；若显式direction_result=纳入（例如Agent实习因招聘类型排除），可保留已知方向等级，但不计候选。
- **所有筛选结果都可独立填写city_priority**，不因方向待核实/排除/不纳入而清空；值应来自run.city_priority_order或“待核实”。规则无城市等级时为""。旧JSON的空城市等级兼容读取并生成非阻断补充提示，新任务应逐行独立评级。
- result=待核实或status=待核实：`uncertainties`或`blocking_issues`至少有一项具体说明。
- city_priority=待核实时，uncertainties/nonblocking_notes/blocking_issues至少一项说明地点缺失。无地域排除策略时，地点缺失通常不阻断；启用city_exclusion_priorities后，未知/空城市等级不能最终纳入，可保留direction_result=纳入及priority、将result设待核实并说明blocking_issues。
- city_priority命中run.city_exclusion_priorities时，result必须为排除，并填写blacklist_reason说明地域原因；即使方向T0也不计候选。同岗多城市仅排除命中的城市行，保留其他有效城市，job_key仍共用。

其他可选字段（建议尽可能依据原文填充，不得猜）：

| 字段 | 含义 |
| --- | --- |
| direction_result | 方向判断，未提供时等于result；仅明确匹配方向才有岗位priority |
| classification_basis | 岗位名称 / 岗位职责 / 招聘专业 / 综合；名称直判要如实填写 |
| classification_confidence | 高 / 中 / 低；不是自动复核开关，名称直判通常为中 |
| blocking_issues | 真实影响最终筛选/当前可投性的阻断事项，字符串；方向已知但最终待核实时必填 |
| nonblocking_notes | 非阻断补充信息，如仅缺完整JD/地点、未确认更高等级等；不逐项自动标红 |
| job_id | 官网岗位/志愿 ID，保留原值；无则空 |
| department | 部门或子单位 |
| batch | 招聘轮次/届次，不同轮次不得混淆 |
| duties | 职责原文或忠实摘要+证据路径 |
| requirements | 任职要求 |
| majors | 招聘专业原文 |
| education | 学历要求 |
| graduation | 届次要求 |
| headcount | 招聘人数原文，文本字段，不参与限投计算 |
| deadline | 官方真实截止日期/招满即止/待核实，不能用推算提醒日期 |
| deadline_basis | 原文、年份/时区等截止判定依据 |
| deadline_source_url | 截止证据链接 |
| application_url | 具体岗位投递链接；需要登录但无权限时仍可保留非敏感公共定位 URL |
| notes | 其他提醒、证据文件相对/绝对路径等 |

`job_key` 去重依赖 Agent 的真实身份识别，不由脚本按标题臆测。若不同城市职责/资格不同，可在多行各记事实；主/城市等级依用户规则，各 result 的独立身份数可能重叠，不得机械相加。

## 5. other_info：补充证据与事项

每项：

```json
{
  "category":"申请须知",
  "company_id":"u1",
  "job_key":"",
  "item":"撤回与重投",
  "value":"是否恢复额度尚未核实；不要撤回试验",
  "source_url":"https://example.invalid/faq",
  "quote":"可填写对应关键原文；没有则留空",
  "checked_at":"2026-10-28T15:30:00+08:00",
  "needs_review":true
}
```

category/item/value 必填非空字符串，其余可选；关联 ID 若填写必须真实存在。checked_at 未提供则按 run.checked_at，只有确实同次核验的事项才可省略。

脚本自动加入：范围和规则元信息、完整性警告、源行忽略原因、统计口径、缺失/复杂限投，以及按真实投递身份和具体问题去重的岗位阻断复核。个人资格未评估只对匹配在招候选按单位汇总为非阻断提示；名称直判或其他补充说明不自动标“需复核”；城市缺失仅在没有硬性地域排除时作为非阻断提示，存在地域黑名单则会影响最终可投性。原uncertainties仍在岗位表完整展示，但非空本身不触发复核。附加证据/流程等由Agent补充，显式other_info.needs_review也应遵循阻断口径，不能机械标全部为true。

## 6. 数量和汇总行为

- 总览行严格按source_companies原表统计项生成；总览公司数不等于内部招聘主体数。岗位总表保留实际子公司，并自动补“原表公司名称”“原表记录ID”用于回查。
- 总岗位数按 `(company_id, job_key)` 去重；多城市行另列数量。岗位与城市评级独立；随附规则城市T2是地域黑名单，命中城市行最终排除且不计候选，已知方向可保留。岗位T2仍是计算机技术方向，不等于地域排除。
- 默认排序先按最终筛选结果（纳入→待核实→不纳入→排除），再按岗位和城市两套顺序辅助展示；评级本身不依赖该排序，不合成单一总等级。
- “筛选匹配在招岗位数”要求 result=纳入、status=在招、eligibility≠不符。
- “其中资格已确认数”进一步要求 eligibility=符合已知条件，仍不保证账户名额或投递成功。
- 理论可投上限仅在单位 coverage=完整时考虑：候选数为0则按筛选记0；存在共享/复杂约束则给文字提醒；额度未知/冲突则给“未核实”；独立明确额度才用 min；明确不限才用候选数。
- 该数始终**未扣历史已投**，并把资格未确认说明放相邻列，不是剩余可投数。
- “全部排除/全部不纳入/全部排除或不纳入”只有非空岗位集合才成立；采集不完整加“已见范围/核验未完成”前缀。空岗位表不产生伪造岗位或全排除标记。
- 不生成跨单位的最大可投总数，防止共享额度重复累计。总览的整体限额不从子公司相加推算；原表公司总范围未完整核验时，不能凭已知子公司完成就给出全量结论。
- Excel按顺序包含四张Sheet：所有单位总览、所有单位岗位总表、其他信息表、筛选后岗位表。第四张仅按job.result为纳入/待核实保留；用户所说“待审核”沿用现有枚举“待核实”，不新增另一种状态。它是总表的精确行子集，表头、内容、多城市粒度和相对顺序保持一致，不重新排序/分类，也不额外筛招聘状态或资格。
- 第四张表不以direction_result、status、eligibility或其他信息表的needs_review作为保留条件；最终排除但方向T0的岗位不保留。匹配筛选结果不等于保证可投，所有原始风险字段保留。空子集仍输出表头；总表及公司总览不受影响。
- 导出返回值filtered_job_city_rows为第四张表的数据行数，不是个人可投数或去重岗位数。companies为总览原表公司行数，audit_subjects为内部核验主体数，unique_positions为全局独立岗位数。无原表记录的指定单位只列其他信息表，不计总览行数。

## 7. 使用与校验

```bash
python "<skill-dir>/scripts/export_excel.py" "<输出目录>/audit.json" -o "<输出目录>/招聘核验.xlsx"
python -m unittest discover -s "<skill-dir>/scripts" -p "test_*.py" -v
```

需要安装依赖时：

```bash
python -m pip install -r "<skill-dir>/requirements.txt"
```

脚本拒绝重复记录、来源无法对账、孤立引用、缺少必要证据、未知额度伪填数值、岗位与单位状态明显冲突、超长单元格及危险 XML 控制字符。长文本改摘要并将原文存证，不要删改事实来绕过校验。

默认拒绝覆盖已有 Excel；仅获授权后用 --overwrite。文件先临时保存、回读校验四张 Sheet 和行列数/无公式，再落到目标路径。导出后仍应抽查单位标记、共享额度、岗位排序、链接与实际证据是否一致。

`assets/example-audit.json` 全是虚构数据，仅用于离线演示和测试。它的规则路径/哈希是明确的占位信息，实际任务必须替换为本次真实来源、规则哈希与时间，不能复制后当作真实结果。
