# 实验报告生成 Skill

根据用户上传的实验报告模板、实验题目与要求、答案代码，分析并生成符合模板的实验报告。
报告中的实现描述、运行结果和结论需要能对应到用户源码或真实证据。

## 功能

- 识别 `.docx` 和旧版 `.doc` 模板的章节、表格、填写区域及排版规则。
- 阅读 Word、PDF、图片或文字形式的实验说明，逐项提取任务与思考题。
- 分析常见语言源码或项目文件夹，建立“实验要求 → 代码 → 模板栏目 → 结果证据”的对应关系。
- 默认交付 `.docx`，按需从最终 Word 文档导出 PDF。
- 保留原始材料；发现缺失、冲突或影响结论的歧义时先询问。
- 区分用户提供的结果、实际运行、静态分析和预期输出，不编造数据、截图或运行经历。

需要具备附件读取、源码分析、文件写入和页面视觉检查能力的 Codex 或兼容 Agent Skills 的环境。
理解材料和撰写正文由模型完成，附带脚本负责结构检查、已确认位置的填充和格式转换。

## 安装

可在支持 `$skill-installer` 的 Codex 中输入：

```text
使用 $skill-installer 安装这个仓库中的 skills/lab-report-generator：
https://github.com/9huchenming1-spec/lab-report-generator-skill
```

也可以克隆或下载仓库，将 `skills/lab-report-generator` 整个目录复制到 Codex 的技能目录。
按 [OpenAI 官方技能说明](https://learn.chatgpt.com/docs/build-skills)，用户级目录为 `~/.agents/skills/`，
项目级目录为 `.agents/skills/`；使用自定义技能目录的环境，请复制到其配置的位置。
保留目录名 `lab-report-generator` 以及内部的 `scripts`、`references` 和 `agents`。

## 使用

上传模板、实验要求和答案代码后，输入：

```text
使用 $lab-report-generator，按照我上传的模板和实验要求，
分析这份答案代码，生成 Word 实验报告。
需要补充材料或遇到影响结果的不明确事项时，先询问我。
```

如需 PDF，补充“同时导出 PDF”。如果已有运行日志、截图或数据，请一并提供。
没有实际结果时，skill 会明确标注预期或待补充；存在缺口的文件称为草稿。

## 环境依赖

| 用途 | 依赖 |
| --- | --- |
| 理解说明、代码和图片 | 支持相应文件与视觉能力的模型环境 |
| 检查与填充 DOCX | Python 3.10+、lxml |
| 提取 PDF 说明文字 | pypdf；扫描页仍需视觉识别 |
| DOC 转 DOCX、DOCX 导出 PDF | Windows 上已安装的 Microsoft Word，或 LibreOffice |
| 检查实际页面 | 环境自带渲染器；也可导出 PDF 后使用 Poppler |
| 运行答案代码 | 该项目自己的解释器、编译器和依赖，按需准备 |

仓库中的 Python 助手依赖：

```sh
python -m pip install -r requirements.txt
```

单独安装 skill 时可执行 `python -m pip install lxml pypdf`。Codex 提供工作区依赖时优先使用其运行环境。
缺少 Word/LibreOffice 时不能直接处理旧版 DOC 或导出 PDF，skill 会说明缺少的工具并询问后续方式。
旧版 DOC 转换后的版式需要与原件对照，不能保证所有专有排版完全一致。

## 演示

`examples/basic/` 是专门编写的脱敏示例：

- `requirements.md`：插入排序和二分查找的实验要求。
- `answer.py`：答案源码。
- `template.docx`：两页 Word 模板。
- `report.docx`：按模板填写的完整演示报告。
- `fill-plan.json`：核对过的填写位置和目标文字。
- `run-results.txt`、`evidence.json`：六组实际执行输入的输出和证据记录。

查看 [生成报告](examples/basic/report.docx) 和 [模板](examples/basic/template.docx)。
本演示使用合成身份，不包含真实姓名、学号或课程材料。

重建演示需要开发依赖，并指定一个没有同名输出文件的目录：

```sh
python -m pip install -r requirements-dev.txt
python tools/build_demo.py --output work/demo
```

这是可复现的演示样例，不能通过更换 `answer.py` 自动生成任意实验报告；处理新材料应调用 skill。

## 文档工具

在 skill 安装目录中执行，或使用脚本的完整路径：

```sh
python scripts/docx_template.py inspect template.docx --output work/template-map.json
python scripts/docx_template.py fill template.docx --plan work/fill-plan.json --output output/report.docx
python scripts/docx_template.py compare template.docx output/report.docx
python scripts/convert_document.py template.doc --output work/template.docx
python scripts/convert_document.py output/report.docx --output output/report.pdf
python scripts/extract_pdf.py requirements.pdf --output work/requirements.txt
```

填写计划必须先分析模板再建立，格式和能力范围见 [文档处理说明](skills/lab-report-generator/references/document-handling.md)。
不要求模板预设占位符；语义明确的空表格段落也可填写。工具对复杂字段、控件和图像段落会拒绝简单文字修改，
这些内容由合适的文档工具局部处理。PDF 提取返回退出码 2 表示有页面必须视觉检查，不表示页面没有实验要求。

## 开发验证

```sh
python -m pip install -r requirements-dev.txt
python tools/validate_skill.py
python -m unittest discover -s tests -v
```

GitHub Actions 检查技能元数据、资源链接和文档工具。Office 转换及最终版式需要在有对应工具的环境中另行验证。
行为场景见 [验收用例](tests/behavioral-cases.md)。

## 目录

```text
skills/lab-report-generator/   可独立安装的 skill
  SKILL.md                    核心工作流程
  agents/openai.yaml          中文界面信息
  references/                 文档处理、证据规则和交付检查
  scripts/                    DOCX、格式转换和 PDF 工具
examples/basic/               公开演示输入与输出
tests/                        工具测试和行为验收场景
tools/                        仓库检查及演示构建工具
```

个人输入与生成文件建议放在已忽略的 `input/`、`output/` 或 `work/`，勿提交真实课程附件、身份信息或凭据。
仓库按作者要求暂不添加许可证。
