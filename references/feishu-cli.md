# 飞书 CLI：授权与只读获取招聘表

本 Skill 支持通过官方 [larksuite/cli](https://github.com/larksuite/cli) 的 `lark-cli` 获取飞书源数据。它是浏览器和直接 API 之外的**独立读取渠道**，不是绕过访问控制的工具，也不是另一个同名社区 `feishu-cli` 的命令别名。

**适用范围**：读取飞书源表、字段、视图和记录。外部招聘网站的当前岗位、详情、限额仍须按主 Skill 浏览核验。CLI 授权不代表招聘网站已登录，更不代表已经完成岗位筛选。

## 1. 渠道选择与安装

用户明确指定 CLI 时采用 CLI；否则可优先复用**已安装且已授权**的 CLI 获取结构化源表，浏览器作为补充和回退。CLI 不可用不阻塞已有可用的只读浏览器/公开分享/API 路径。

先确认本机版本及真实命令，不猜 flag：

```bash
lark-cli --version
lark-cli --help
lark-cli base --help
```

未安装时，将官方安装方式告知用户，**获得安装同意后**再执行；仅要求修改此 Skill 不等于授权安装、创建应用或登录。需 Node.js/npm/npx，支持版本以官方包的 engines 为准：

```bash
npx @larksuite/cli@latest install
```

安装可能写入 CLI 及配套官方 Skills，不应静默更改现有全局工具。官方 `lark-shared`、`lark-base` 已可用时先读取它们，再按命中场景读取相关 reference。普通电子表格改用 `lark-sheets`，不要当成 Base。

以下命令依据官方文档整理，实际调用前必须用**本机该版本** `--help` 和对应官方 reference 核对。若本机版本无相应命令或 flag，按已确认的兼容路径读取，或回退浏览器；不能凭空补命令。

## 2. 权限从哪里来

CLI 通过飞书应用与 OAuth 获取授权，和 `ego-browser` 的 Cookie/登录态相互独立，**不会自动继承浏览器权限**。读取需要同时满足：应用已开通对应能力、用户授予所需 scope、目标文档/数据行字段对该身份可见；租户策略和高级权限可能另外限制。

本任务默认显式使用 **`--as user`**。用户没有明确指定并同意应用身份时，不因读取失败或空结果就切换为 `--as bot`；应用身份有独立的 scope 与资源授权。

### 已有配置与授权

```bash
lark-cli auth status --json --verify
```

只摘取是否已登录、是否验证有效、当前身份及必要的权限摘要。完整输出可能含姓名、open_id 等个人标识，不能直接贴进公开仓库、Excel 或持久公共日志。禁止运行或转存任何“显示 token/secret”的调试输出。

### 首次配置（仅用户同意后）

应用尚未配置时使用官方交互流程，不要求用户把 app secret 发到聊天或写进项目：

```bash
lark-cli config init --new
```

此命令可能创建/配置应用并等待浏览器确认，不属于读取数据的必要默认动作。复用已有应用时按本机官方配置指南选择，不覆盖现有配置、不擅自新建应用。

### 最小范围 OAuth（仅实际需要连接时）

根据当前命令的官方权限要求、应用可用 scope 或错误中的 missing_scopes，确认所需的**最小只读 scope**后发起。不同 Base/Sheets/API 版本所需 scope 可能不同，不能把旧 bitable 权限名硬套到新 Base 命令，也不要把所有缺失候选 scope 当作必须全部授予。

```bash
lark-cli auth login --scope "<已确认所需的只读scope>" --no-wait --json
```

本 Skill 不默认使用全域或推荐大范围授权，也不请求消息、通讯录等无关能力；不为了自动通过而授予写权限或改变文档 ACL。权限不足时区分 scope、资源 ACL、租户管控，不能一律反复登录。

Agent 代理授权采用分轮流程，避免阻塞用户看到链接：

1. 从本次返回中提取 verification_url/device_code，只作为本次短时私有会话数据处理。
2. 原样向用户展示授权 URL；按官方共享规范生成二维码，先链接后二维码。文件放在**仓库和全局 Skill 之外的本次私有工作目录**，从该目录运行，路径用 cwd 下相对路径：

   ```bash
   lark-cli auth qrcode "<本次返回的授权URL>" --output "auth-qr.png"
   ```

3. 结束当前轮，提示用户完成授权后回复；不要展示链接后立刻阻塞轮询。
4. 用户确认已授权后，在后续轮完成本次登录，再核验登录状态：

   ```bash
   lark-cli auth login --device-code "<本次device_code>"
   lark-cli auth status --json --verify
   ```

授权链接、二维码、device_code、refresh token 和应用配置不能写入 audit.json/其他信息表、仓库或全局 Skill；不要复用过期链接。若 CLI 提示未获得 refresh_token，应说明其到期/无法自动续期及所需管理员操作，不自动循环重新授权。注销本机登录不等于撤销服务端授权，后者由用户在飞书授权管理中处理。

## 3. 解析真实表、视图与字段（只读）

先解析用户给出的链接，避免把 wiki/share token、BaseApp token 或其他顶层块当作数据表：

```bash
lark-cli base +url-resolve --url "<用户提供的飞书链接>" --as user
```

使用实际返回的 base_token、table_id、view_id 和 resource_type/block_type；只有目标确为 table 才进入记录查询。需要消歧时只读获取真实目录或视图列表，不靠同名猜选。

```bash
lark-cli base +field-list --base-token "<base_token>" --table-id "<table_id>" --as user
lark-cli base +view-get --base-token "<base_token>" --table-id "<table_id>" --view-id "<view_id>" --as user
lark-cli base +view-get-filter --base-token "<base_token>" --table-id "<table_id>" --view-id "<view_id>" --as user
```

按真实字段名/ID投影单位、批次/届次、投递入口、公告、截止时间等必需列，尽量不取在投人员、手机号等与任务无关字段。字段隐藏不等于没有数据，但不能突破行/列访问权限。

**view_id 出现在 URL 或请求里不是范围已验证的充分证据。** 检查返回的 query_context、实际视图筛选和必要的 UI/记录 ID 对照，再设置 run.scope_verified。

## 4. 按指定视图分页读取

先读取本机 `lark-cli base +record-list --help`。下例适用于官方支持 --view-id/--offset/NDJSON manifest 的版本。占位符必须换为实际值，--field-id 可重复使用字段 ID 或该版本明确支持的字段名；从**私有工作目录**运行，输出不要落在仓库内：

```bash
lark-cli base +record-list --base-token "<base_token>" --table-id "<table_id>" --view-id "<view_id>" --field-id "<单位字段ID>" --field-id "<批次字段ID>" --field-id "<投递链接字段ID>" --limit 2000 --offset 0 --format ndjson --output "records-000.ndjson" --as user
```

不要只投影示例的三列：依据实际任务继续添加公告/截止时间等已核实字段。输出包括 NDJSON 记录和同名 manifest；按本机版本确认文件名和 stdout 摘要结构，不自行臆造字段。官方当前协议的关键字段是：

- records_count、has_more、next_offset：本块数量、是否还有记录、下一块起点。
- query_context：实际 record_scope、view/filter/sort/字段投影；指定视图预期为 `view_filtered_records`。
- rev：来源快照版本，多块之间需一致。
- columns：field_id → 实际列名/类型映射，避免列改名或链接展示值丢失。

has_more=true 时仅使用返回的 next_offset 继续，固定表、视图、筛选、排序及投影，改用不同输出文件；直到 has_more=false。不要假设第一块就是全部、不要把2000条上限视为整表数量，也不要把其他命令的通用 --page-all 无条件套到本命令。空页仍宣称 has_more、游标重复、record_id重复或rev改变时停止作完整结论，重读或标部分完成。

### 视图与指定单位交集：重要陷阱

官方 Record SOP 明确：**--filter-json 会覆盖视图筛选，而不是自动取交集**；--sort-json 同样覆盖视图排序。因此本 Skill 默认先保留原视图范围完整读取，再在本地精确选择用户指定单位。若要服务端过滤，必须读懂原条件并显式构造交集、核对返回范围；无法复现时回退原视图读取，不修改线上 view filter。

只指定单位而没有指定视图时，可在已确认的整表范围查询/搜索并消歧，但不能把搜索零结果当停招。+data-query 的聚合或去重维度行不包含完整逐条原始记录，不能取代本任务的单位来源台账。

### 链接和成功响应

- NDJSON 中的链接可能为 URL、Markdown 链接或展示文本；核实真实 href。若导出只有显示文字，按已确认的详情/字段接口补取，不能把显示文字拼成链接。
- CLI 的 JSON 成功信封使用 `ok == true` / 退出码0，不检查不存在的顶层 code=0；错误在 stderr，检查 error/missing_scopes。
- 直接官方 API 的原始 code/msg 信封与 CLI 信封不同；不同渠道分别解析，不能混用。

## 5. 普通飞书电子表格

确认对象是 Sheets 后加载官方 lark-sheets 的读取 reference，并先查真实工作簿/子表：

```bash
lark-cli sheets --help
lark-cli sheets +workbook-info --url "<电子表格链接>" --as user
```

读取区域按本机版本的 +csv-get/+cells-get 等**只读**命令执行，记录工作表、行范围/稳定行键及现有筛选视图；不要用 Base 的 table_id/view_id 代替 Sheets 参数。已有筛选不能为了获取数据而改写；CSV 可能丢富文本 href 时用可恢复链接的只读接口补充。

## 6. 接入审计结果及失败回退

- 将已验证的 record_id/稳定行键填入 run.source_record_ids，按主 Skill 归并或展开到 companies.source_record_ids；说明行附理由忽略。
- 在 run.notes 或 other_info 记录“采集渠道=lark-cli”、CLI版本、身份类型 user/bot、范围/分页/版本一致性摘要和非敏感本地证据路径；不记录姓名、open_id、授权链接或凭证。
- 原始 NDJSON、manifest、字段/视图响应均可能含私有表标识和业务数据，只存私有工作目录，不提交到 Git 或复制进全局 Skill。
- 访问失败、缺scope、资源ACL不允许、版本不支持、不能证明视图过滤等，标明具体原因；可用时回退用户已有权限的浏览器或其他只读渠道，不切身份、不扩大范围、不把失败判为停招。
- 不能借 CLI 自动创建/更新/删除记录、视图、分享或权限；官方工具具有写能力不代表本 Skill 允许写入，--dry-run 也不是擅自实施写操作的授权。

## 官方依据

- [官方 CLI README：安装与认证](https://github.com/larksuite/cli/blob/main/README.md)
- [官方共享规范：身份与权限](https://github.com/larksuite/cli/blob/main/skills/lark-shared/references/lark-shared-identity-and-permissions.md)
- [官方 Base Skill](https://github.com/larksuite/cli/blob/main/skills/lark-base/SKILL.md)
- [官方 Record SOP：视图范围、分页与 manifest](https://github.com/larksuite/cli/blob/main/skills/lark-base/references/lark-base-record-query-and-analysis-sop.md)

核对时官方 npm 包版本为1.0.99，参考仓库修订为05ac58ffa089fe7a171205ac1d1e97ee50196ec8。这不是版本锁或已完成真实表访问的证明；运行时以用户本机版本和当前官方契约为准。
