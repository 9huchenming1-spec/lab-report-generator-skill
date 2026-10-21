# 文档处理

本 skill 的路径均相对自己的安装目录。以下命令中的 `python` 指当前环境可用的 Python 3.10+；
Codex 提供工作区依赖加载工具时优先使用其返回的 Python。公开仓库根目录的 `requirements.txt` 列出助手工具依赖；
单独复制 skill 时可执行 `python -m pip install lxml pypdf`。没有运行或安装权限时不要擅自改变环境。

## DOCX 检查与定点填充

```sh
python scripts/docx_template.py inspect path/to/template.docx --output work/template-map.json
python scripts/docx_template.py fill path/to/template.docx --plan work/fill-plan.json --output output/report.docx
python scripts/docx_template.py compare path/to/template.docx output/report.docx
```

inspect 输出正文、表格、页眉页脚、脚注/尾注中的段落，带包部件名、结构路径、原文及是否可安全编辑；
另列包部件哈希、节设置、表格、图片数量、字段、控件和修订标志。文本框段落亦会被列出，但须视觉核对。
脚本支持常见 Transitional OOXML；不支持 Strict OOXML 时会明确报错。

填充计划：

```json
{
  "template_sha256": "填写 inspect 实际输出的哈希",
  "operations": [
    {
      "part": "word/document.xml",
      "path": "/*[1]/*[1]/*[2]",
      "expected": "姓名：{{姓名}}",
      "replace": {"old": "{{姓名}}", "new": "演示同学"}
    },
    {
      "part": "word/document.xml",
      "path": "/*[1]/*[1]/*[4]/*[2]/*[1]/*[2]",
      "expected": "",
      "text": "这里填写已经分析和核对过的答案。\n这里是同一段中的第二行。"
    }
  ]
}
```

路径必须来自本次 inspect，不能照抄示例。每个段落最多一个 operation；需要多处替换时先计算整段目标文字再用 text。
`replace` 要求 old 在该段恰好出现一次；支持占位符被 Word 拆成多个 run 的情况，保留周围文字和其样式。
`text` 将整段文字设置为目标值，沿用该段第一个 run 的文字格式；多种字体的混排段落优先用 replace。
段落样式、表格、分页和其他包部件不重建。换行和制表符写成 Word 的 break/tab 节点。
遇到字段、超链接、绘图、书签、批注、修订或控件等复杂段落，工具拒绝填充，不为“成功”删除这些结构。

脚本检查原件哈希、路径及 expected，全部通过才写出新文件；拒绝覆盖输入和已有输出。
compare 列出新增、删除和内容变化的包部件，用于确认未修改样式、页眉页脚、图片等保留项。
compare 不是语义判断或视觉验证；仍需逐页检查。

## 普通模板与复杂排版

模板不需要事先含 `{{占位符}}`。定位语义明确的空段落、表格单元格或示例答案段，确认可改后用路径填充。
不能自动将所有空段落视为答案区。需要多个真实段落、完整代码、图片、公式或增行表格时，使用环境已有的文档工具局部编辑工作副本。
优先继承原段落/单元格样式，保留合并单元格、页码、样式定义及关系。编辑后比较包部件和模板合同。
受保护模板、数字签名、复杂控件或无法保留的专有排版，先明确说明并询问处理方式。

## DOC 转换及 PDF 导出

```sh
python scripts/convert_document.py input/template.doc --output work/template.docx
python scripts/convert_document.py output/report.docx --output output/report.pdf
```

支持已安装的 Microsoft Word（Windows）或 LibreOffice。自动模式在 Windows 优先 Word，其他平台使用 LibreOffice；
可用 `--engine word` 或 `--engine libreoffice` 指定，不在失败后静默切换引擎。路径不在 PATH 时用 `--soffice` 指定 LibreOffice 可执行文件。
Word 通过独立的隐藏 COM 实例打开只读工作副本，禁用宏、链接自动更新和交互提示；退出时只关闭自己的实例。
LibreOffice 使用独立用户配置目录。工具不安装转换器，不覆盖原件/已有输出，失败时不留下最终输出。

旧版 `.doc` 转换会改变文件结构，需对照原件核对页面、字体、表格和栏目，不能承诺像素级一致。
没有转换器时，询问是否安装转换器或请用户另存为 DOCX；不要改后缀，不要把纯文本提取当作模板转换。
按需 PDF 必须从最终 DOCX 导出。Word/LibreOffice 的目录和交叉引用可能需要更新；导出后核对页码与目录。

## PDF 与图片说明

```sh
python scripts/extract_pdf.py input/requirements.pdf --output work/requirements.txt
```

每页附页码。页面缺少有效文本会标记 `NEEDS_VISUAL_REVIEW` 并返回退出码 2：此时仍保留已提取内容，
必须渲染/查看相应页面，不能忽略扫描页。数学公式、表格、多栏布局即使有文本，也要按需要查看原页。
图片直接用可用视觉工具识别；OCR 只是辅助，需要核对题号、符号、上下标和数值。
模糊、遮挡或无法访问的部分向用户说明并请求清晰材料。

## 渲染与版式检查

使用当前环境已提供的文档/PDF 渲染器。Codex 的 documents skill 可用时，遵循其中的模板处理与渲染规则。
通用替代是先从工作副本导出 PDF，再使用 Poppler `pdftoppm -png` 生成页面图像，逐页查看；这不是运行代码的截图。
渲染中间件只用于检查，不覆盖 DOCX。记录引擎和缺失字体，字体替换造成的分页变化需说明并解决。
无渲染器或视觉能力时，询问用户接受何种检查方式，并准确说明仅完成结构检查。
