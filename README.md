# recruitment-link-audit

面向 AI 编码助手的可复用 Skill：核验飞书招聘表中的指定视图或单位，检查投递链接是否仍有本轮招聘岗位，按**用户每次指定的规则文件**筛选，并导出三张 Sheet 的 Excel。

这是一个 **Agent Skill + 离线导出器**，不是一键爬虫。网页访问、证据核验和语义筛选由加载 Skill 的助手执行；Python 脚本只校验结构化结果并生成 Excel。

## 功能

- 通过官方飞书 CLI（`lark-cli`）、浏览器或已有授权的只读 API 获取源表。
- 限定飞书视图、指定单位或二者交集，保留源记录追踪。
- 区分在招、已停止、暂未开放、当前无岗位与无法核实。
- 读取用户指定的规则，不内置个人岗位偏好或私有画像。
- 明确岗位名称可直接判断，职责有反证则优先；非技术岗未见计算机专业依据不再一律悬置。
- 保留全部结果，区分方向、最终筛选与个人资格；只把真正阻断项标需复核，个人资料缺失按单位提示。
- 岗位与城市独立评级，排除/方向待核实的岗位也保留已知城市等级。
- 多城市按真实独立投递身份去重，不将招聘人数当成可投岗位数。
- 核实官方限投数，区分单位独立限额与集团共享/复杂分层约束。
- 未核实的限制不默认填 1、2 或“不限”，不把理论上限伪装成个人剩余名额。
- 默认只读，不修改飞书、不创建或提交求职申请。

## Excel 内容

| Sheet | 内容 |
| --- | --- |
| 所有单位总览 | 所有目标单位、招聘状态、投递总结、独立岗位与匹配岗位数、官方最多可投岗位数、理论可投上限及依据；显式标记全部排除/全部不纳入/混合情况 |
| 所有单位岗位总表 | 全部已见的范围内岗位，包括排除、不纳入、停招及待核实岗位；含岗位/独立城市优先级、判断依据类型/置信度、阻断与补充信息 |
| 其他信息表 | 本次范围、规则哈希、统计口径、异常、共享额度、申请须知与复核事项 |

“按筛选的理论可投上限”未扣个人历史已投，资格未确认时也会提示。复杂共享约束不直接求和。具体口径见 [核验细则](references/verification.md) 和 [数据契约](references/data-contract.md)。

## 安装为全局 Skill

支持 `~/.agents/skills/` 的 Agent 可将本项目放入其下。以首次安装、尚无同名目录为例：

```bash
git clone https://github.com/Li7777777/recruitment-link-audit.git ~/.agents/skills/recruitment-link-audit
```

Windows 对应位置是用户主目录下的 `.agents/skills/recruitment-link-audit`。也可只复制 `SKILL.md`、`references/`、`scripts/`、`assets/` 和 `requirements.txt`；如需随附的公开规则，一并复制 `岗位优先级.md`。已有同名 Skill 时先比较内容、备份，再更新，避免覆盖自己的改动。

Pi 用户安装后执行 `/reload`；其他助手按各自的 Skill 加载方式刷新。

## 使用示例

向助手提供飞书链接、目标视图或单位，以及自己的规则文件：

```text
使用 recruitment-link-audit 检查 <飞书视图链接>，
按 @岗位优先级.md 筛选，核实每家最多能投几个岗位，
输出 Excel 到我指定的目录。
```

在 Pi 中也可以显式调用：

```text
/skill:recruitment-link-audit 检查 <飞书链接> 中的 <单位列表>，规则用 @<规则文件路径>，导出 Excel。
```

飞书源表可优先复用已安装且已授权的官方 `lark-cli`；用户指定其他渠道时遵从用户选择。外部招聘网站仍需要浏览器，优先使用 `ego-browser`，先加载其 Skill 并遵循本机指南。**本仓库不提供浏览器登录态、CLI 授权、飞书凭证或任何账号密码。**

### 使用官方飞书 CLI 获取源表

官方项目：[larksuite/cli](https://github.com/larksuite/cli)，npm 包 `@larksuite/cli`，命令 `lark-cli`。不要与其他同名社区工具混用。未安装时可在同意安装后按官方流程执行：

```bash
npx @larksuite/cli@latest install
lark-cli --version
lark-cli auth status --json --verify
```

已有登录可直接复用；首次应用配置与 OAuth 需用户确认，按当前版本选择最小只读 scope，不默认申请全域/写权限。读取始终明确用户身份，检查真实视图范围并完整分页。CLI 不自动继承浏览器登录态，也不赋予原本没有的文档访问权限。

安装、授权分轮确认、视图读取命令、`--filter-json` 覆盖视图的陷阱、NDJSON/manifest 完整性及回退方式，见 **[飞书 CLI 读取指南](references/feishu-cli.md)**。本次接入不包含自动安装、真实授权或私有表访问测试。

### 可选的公开岗位规则

经用户明确授权，仓库随附 [岗位优先级.md](岗位优先级.md)，保存本次同意公开的规则版本：仅校招正式岗位、岗位T0～T3、独立城市T0/T1/T2及排除规则。城市T0为江浙沪，T1为其余明确地区；城市T2（新疆、西藏、内蒙古、山西、北京、青海）是硬性工作地点排除，不设置城市T3。该文件不含个人身份、联系方式或访问凭证。

这是**可显式选用的规则文件，不是 Skill 强制默认值**。每次仍以用户指定文件为准；要使用这一版，请明确写“按仓库中的 @岗位优先级.md 筛选”。修改自己的私有规则时请另存到仓库外，不要直接改动并公开提交。

名称直判、缺专业的处理、官方校招上下文继承和非阻断复核口径见 [判断策略](references/decision-policy.md)。方向T0也不能豁免城市T2地域排除；同岗多可选城市仅排除命中地区的城市行，地点不明时保留方向、阻断最终纳入；旧报告不会被脚本按标题静默改写，需重跑时使用新规则和原始证据。

其他未经公开授权的个人规则、CLI 配置/授权链接/二维码/令牌、实际招聘数据，请存到独立的本地工作目录，不要提交到公开仓库。

## 离线导出与测试

需要 Python 3.10+。建议在虚拟环境中安装依赖：

```bash
python -m venv .venv
# 激活虚拟环境后：
python -m pip install -r requirements.txt
python -m unittest discover -s scripts -p "test_*.py" -v
```

演示导出（全部为虚构数据，不会访问招聘网站）：

```bash
python scripts/export_excel.py assets/example-audit.json -o outputs/example.xlsx
```

真实核验先按 [数据契约](references/data-contract.md) 生成本地 `audit.json`，再导出：

```bash
python scripts/export_excel.py <本地工作目录>/audit.json -o <本地工作目录>/招聘核验.xlsx
```

默认拒绝覆盖已有 Excel；只有明确需要覆盖时才使用 `--overwrite`。脚本包含来源对账、证据字段校验、规则文件可访问时的哈希校验、公式注入防护和 Excel 回读检查。

## 项目结构

```text
SKILL.md                        # Skill 入口和完整执行流程
岗位优先级.md                    # 用户授权公开、可显式选用的规则版本
references/verification.md      # 状态/范围/限额核验口径
references/decision-policy.md   # 名称直判、独立评级和阻断复核分流
references/feishu-cli.md        # 官方 CLI 安装、授权与视图只读获取
references/data-contract.md     # 审计 JSON 字段与导出约定
scripts/export_excel.py         # 离线导出器
scripts/test_export_excel.py    # 离线导出回归测试
scripts/test_cli_docs.py        # CLI 指南与只读示例约束检查
scripts/test_decision_policy.py # 公开规则与独立城市分级约束检查
assets/example-audit.json       # 明确标注的虚构示例
requirements.txt               # openpyxl
SECURITY.md                     # 隐私与提交检查说明
```

## 隐私与安全

仓库包含通用流程、代码、虚构示例，以及用户明确授权公开的根目录 `岗位优先级.md`；不包含真实招聘表、个人简历、未授权公开的私有规则、飞书应用凭证、API Key 或导出报告。`.gitignore` 仅为该根目录规则开了例外，运行数据和其他敏感文件仍应排除，提交前仍需检查。

可使用 [Gitleaks](https://github.com/gitleaks/gitleaks) 在本地扫描：

```bash
gitleaks dir . --redact --no-banner
gitleaks git . --redact --no-banner
```

更多边界与注意事项见 [SECURITY.md](SECURITY.md)。扫描无告警不等于绝对不存在隐私风险；公开提交前仍需核对待提交文件和 Git 历史。
