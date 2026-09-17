# Interview Minutes Skill

将中文访谈转写稿整理为机构级 Word 纪要的 Agent Skill。仓库内置完整的企业江滨格式模板，并通过强制生成器和格式守门器保证模板不会被通用 LLM 文档样式覆盖。本仓库含内部完整样例，必须保持为私有仓库。

## 核心能力

- 行业术语和音近公司名纠错
- 国内企业“企查查 MCP＋官网／权威来源”双重核验
- 事实、数字和判断逐条回溯原始转录
- 企业 Word 模板 OOXML 级继承
- 标题、文件名和元信息标准化
- 内容与版式双重独立审核
- 模板指纹、字体、编号和分页自动拦截
- PDF 逐页视觉检查

普通 LLM 往往会重新选择字体、手写编号或按主题随意命名文件。本 Skill 的生成器只克隆模板原生属性；任何字体覆盖、模板部件变化、重复编号、错误文件名或占位元信息都会触发 `FORMAT CHECK FAILED`，禁止交付。

## 目录

```text
interview-minutes-skill/
├── SKILL.md
├── agents/openai.yaml
├── assets/
│   ├── build_minutes.py
│   ├── glossary_pcb_ldi.md
│   ├── review_checklist.md
│   ├── template_manifest.json
│   └── template_minutes.docx
├── references/
│   ├── content_protocol.md
│   ├── format_protocol.md
│   └── name_verification.md
├── scripts/
│   ├── qcc_mcp_call.py
│   └── verify_docx_format.py
└── tests/test_format_guard.py
```

模板保留完整真实样例及其原生 OOXML，用于稳定复刻企业版式、字体、间距、编号、粗体范围和页面结构。不得将模板或仓库转为公开可见。

## 安装

> 必须克隆并安装整个仓库，不能只复制 `SKILL.md`。企业模板、生成器、模板指纹和格式守门器缺一不可；资源不完整时应停止运行，而不是降级为普通 LLM 文档生成。

### Codex

```bash
git clone https://github.com/kppzhouau-beep/interview-minutes-skill.git
mkdir -p ~/.codex/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.codex/skills/interview-minutes
```

团队使用共享 Skill 目录时：

```bash
mkdir -p ~/.agents/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.agents/skills/interview-minutes
```

### Claude Code

```bash
mkdir -p ~/.claude/skills
ln -s "$(pwd)/interview-minutes-skill" ~/.claude/skills/interview-minutes
```

重新启动客户端后，通过 `$interview-minutes` 显式调用；支持自动发现的环境也可在上传访谈转写稿时自动触发。

## 依赖

```bash
python3 -m pip install -r requirements.txt
```

逐页视觉审核还需要 LibreOffice 和 Poppler。若运行环境内置文档渲染工具，优先使用内置版本。

## 使用方式

向 Agent 提供逐字稿、项目名称、访谈对象、访谈类型、日期和元信息：

```text
$interview-minutes
请将这份专家访谈逐字稿整理成高标准 Word 纪要。
项目名称：示例项目
访谈对象：某机构
访谈日期：2026年9月11日
访谈地点：线上
访谈人员：张三、李四
严格使用随 Skill 提供的企业模板，不得编造，并完成内容与版式双重审核。
```

Skill 会生成内容 JSON，并强制调用：

```bash
python3 assets/build_minutes.py \
  --payload /absolute/path/payload.json \
  --output-dir /absolute/path/output
```

成功输出必须同时包含：

```text
FORMAT CHECK PASSED
SAVED /absolute/path/项目名称_访谈对象专家访谈纪要_YYYYMMDD.docx
```

如果用户提供另一份参考 Word，在命令中增加 `--reference /absolute/path/reference.docx`。

## 独立格式检查

```bash
python3 scripts/verify_docx_format.py /absolute/path/output.docx \
  --reference assets/template_minutes.docx \
  --project "示例项目" \
  --subject "某机构" \
  --interview-type "专家访谈" \
  --date-compact 20260911
```

校验器检查关键模板部件、模板外字体、段落和 run 属性、文件名、标题、元信息、重复手写编号、页面设置和受控分页。

## 企业名称核验

企查查 MCP 不是安装本 Skill 的前提，但团队环境具有合法可用的企查查 MCP 时，国内企业名称核验必须优先调用，并与官网或权威来源交叉验证。在自己的 Agent 环境中配置相关 Server 和 Token；不要把 Token 写入仓库、脚本或纪要。

原生工具发现异常时，必须先使用以下直连发现路径，不得仅因当前会话显示 0 个 QCC 工具就跳过企查查：

```bash
python3 scripts/qcc_mcp_call.py --server qcc-company --list-tools
```

脚本会从用户配置读取 Server，并从环境变量读取 Token，不会打印 Token。直连路径同一请求最多重试 1 次。未配置企查查或两条调用路径均失败时，再按 `references/name_verification.md` 使用其他合法工商数据库及权威一手来源核验，不得伪称已调用企查查。

## 隐私与合规

- 本私有仓库含完整企业模板和内部样例；授权范围仅限内部团队，不得转为公开仓库或向无关第三方分发。
- 新增原始逐字稿、客户名单、个人信息和 Token 不应提交到仓库。
- 外部检索仅用于名称核错，不得把检索得到的新业务事实写进单场访谈纪要。
- Skill 在本地生成 Word；是否调用外部名称核验服务由使用者自己的配置和权限决定。
- 处理尽调材料时，应遵守所在机构的数据和保密政策。

## 回归测试

```bash
python3 -m unittest discover -s tests -v
```

测试覆盖正确模板输出、字体／样式篡改、错误文件名、重复手写编号和占位元信息。

## License

[MIT](LICENSE)
