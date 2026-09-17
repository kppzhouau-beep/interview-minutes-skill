# Interview Minutes Skill

一个面向 Codex／Claude Code 等 Agent 运行环境的中文访谈纪要 Skill。它将带时间戳和说话人标签的语音转写稿整理为机构级 Word 纪要，并执行：

- 行业术语与音近公司名纠错
- 企业／机构名称双重核验
- 事实、数字和判断逐条回溯原始转录
- 用户指定 Word 模板的 OOXML 级复刻
- 内容与版式双重独立审核
- 逐页 PDF 视觉检查

## 目录

```text
interview-minutes-skill/
├── SKILL.md
├── agents/openai.yaml
├── assets/
│   ├── build_minutes.py
│   ├── glossary_pcb_ldi.md
│   ├── review_checklist.md
│   └── template_minutes.docx
├── references/
│   ├── format_protocol.md
│   └── name_verification.md
└── scripts/qcc_mcp_call.py
```

仓库中的 Word 模板已去除真实访谈内容、姓名及项目信息，仅保留版式、编号和样式定义。

## 安装

### Codex

```bash
git clone https://github.com/kppzhouau-beep/interview-minutes-skill.git
mkdir -p ~/.codex/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.codex/skills/interview-minutes
```

如果团队使用共享的 `~/.agents/skills` 目录，将最后两行改为：

```bash
mkdir -p ~/.agents/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.agents/skills/interview-minutes
```

### Claude Code

```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.claude/skills/interview-minutes
```

重新启动对应客户端后，即可通过 `$interview-minutes` 显式调用；支持 Skill 自动发现的环境也可在上传访谈转写稿时自动触发。

## Python 依赖

```bash
python3 -m pip install -r requirements.txt
```

版式审核还建议安装：

- LibreOffice：把 `.docx` 转为 PDF
- Poppler：使用 `pdftoppm` 逐页渲染检查

## 使用示例

```text
$interview-minutes
请将这份专家访谈逐字稿整理成高标准 Word 纪要。
访谈日期：9月11日
访谈地点：线上
访谈人员：张三、李四
格式严格参考我提供的模板；不得编造，并完成内容与版式双重审核。
```

## 企业名称核验

企查查 MCP 是可选能力，不是安装本 Skill 的前提。团队具有合法企查查权限时，可在 Codex MCP 配置中设置相关 Server 和 Token 环境变量；不要把 Token 写入仓库、脚本或纪要。

当原生工具发现异常时，`scripts/qcc_mcp_call.py` 可执行 MCP 工具发现：

```bash
python3 scripts/qcc_mcp_call.py --server qcc-company --list-tools
```

未配置企查查时，按 `references/name_verification.md` 使用其他合法工商数据库及权威一手来源完成核验。

## 隐私与合规

- 不要把原始访谈、客户名单、个人信息、Token 或内部模板直接提交到公开仓库。
- 外部检索只用于名称核错，不得把检索到的新业务事实写进单场访谈纪要。
- 处理敏感尽调材料时，应遵守所在机构的数据和保密政策。

## License

[MIT](LICENSE)
