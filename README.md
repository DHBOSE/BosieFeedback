# 课堂反馈生成器

面向教培机构 / 独立老师的**课堂反馈报告一键生成工具**：粘贴「课堂速览 + 课堂纪要」，AI 自动归纳成结构化课堂报告，产出排版精美的 **Word / PDF**，并附知识点脑图与家长版文字反馈。

纯本地运行，数据不出本机；支持 DeepSeek / Kimi / 阿里云百炼 / 任意 OpenAI 兼容端点。

---

## ✨ 功能一览

| 模块 | 说明 |
|------|------|
| 📄 报告生成 | 速览 + 纪要 → 结构化报告（封面、课堂内容、课堂表现、纪要、收获、作业、易错分析、学习建议），封面五要素可手动指定 |
| 🧠 AI 脑图 | 未上传脑图时自动生成 Mermaid 知识点脑图：竖版长图、圆角矩形、层级深度按内容自适应（2~5 级）、按分支智能切片分页，矢量嵌入 Word（SVG），打印清晰 |
| 💬 家长反馈 | 报告生成后自动生成微信风格简易文字反馈（年级科目 / 授课内容 / 整体表现 / 学习建议 / 课后作业），可编辑、一键复制 |
| ✏️ 在线修改 | 生成后所有板块可直接编辑，「应用修改并刷新预览」重新排版 |
| 👀 预览与打印 | Word 内嵌预览（docx-preview）/ PDF 预览切换，一键导出或直接打印 |
| 🏫 机构品牌 | 封面 / 页眉 / 页脚分别上传 Logo（PNG 透明底 / JPG）、机构名、联系方式；页码从正文起算；**正文大水印**（文字或图片，大小 / 透明度 / 角度 / 23 种中文字体 / 颜色均可调）；品牌全局生效，旧报告重新导出自动用最新设置 |
| 📚 学生档案 | 按学生姓名自动归档，跨报告学情趋势（进步轨迹 / 课堂亮点 / 待改进 / 教学建议，可限定课程范围），学情 PDF 导出 |
| 📦 批量生成 | 多节课顺序生成，每节可单独上传脑图、单独填写封面信息；完成后打包 ZIP（docx ± pdf ± 家长反馈 txt） |
| 🔌 多模型 | DeepSeek / Kimi / 百炼 / 自定义端点，顶栏模型快切，Key 仅存本地且掩码显示 |
| 🕘 历史报告 | 自动保存每次生成，随时载入重新编辑 / 导出 / 删除 |

## 🖥 本机环境要求（正常使用前请确认）

| 需要的组件 | 必需性 | 用途 | 没有它会怎样 |
|------------|--------|------|--------------|
| **Windows 10 / 11** | ✅ 必需 | 启动器（VBS）、PDF 转换、中文字体均依赖 Windows | 无法直接运行 |
| **Python 3.9+** | ✅ 必需 | 运行后端服务 | 无法启动 |
| **python-docx、Pillow** | ✅ 必需 | `pip install python-docx Pillow`；生成 Word、烘焙水印图片 | 无法启动 |
| **LLM API Key**（DeepSeek / Kimi / 百炼 / 任意 OpenAI 兼容端点） | ✅ 必需 | AI 归纳报告、脑图、家长反馈（需联网调用对应 API） | 无法生成内容 |
| **Microsoft Edge** | ⚠️ 建议 | 无头渲染 AI 脑图（Win10/11 系统自带，无需另装） | 报告正常生成，但脑图位置留空白占位 |
| **Microsoft Word** 或 **LibreOffice** | ⚠️ 建议 | docx → PDF 转换（优先 Word，保真度高；LibreOffice 兜底） | Word 报告正常导出，仅无法转 PDF |
| **常见中文字体**（微软雅黑/宋体/仿宋/楷体等，`C:\Windows\Fonts` 系统自带） | ⚪ 可选 | 正文文字水印的字体选择 | 水印自动回退到微软雅黑 |

**不需要**安装 Node.js / npm——前端全部内置于 `static/`，`package.json` 里的 `npm run dev` 只是 `python server.py` 的别名。

## 🚀 快速开始

```bash
pip install python-docx Pillow

cd webapp

# 复制配置模板并填入你的 API Key（也可启动后在网页「设置」里填）
cp config.example.json config.json

python server.py --host 127.0.0.1 --port 7100
# 打开 http://localhost:7100/
```

Windows 用户也可直接双击 `webapp/启动课堂反馈生成器.vbs`（无窗口后台启动 + 自动打开浏览器），`webapp/停止课堂反馈生成器.vbs` 停止。

## 📝 使用流程

1. 填写封面信息（课程标题 / 学科年级 / 授课形式 / 日期，可不填由 AI 归纳）与学生姓名（必填，用于归档）
2. 粘贴「课堂速览」（带时间戳的分段小结）与「课堂纪要」
3. 可选：上传自己的脑图；不传则由 AI 自动生成
4. 点「生成课堂反馈」→ 左侧修改、右侧预览 → 导出 Word / PDF / 打印
5. 底部「家长反馈 · 简易版」可修改后一键复制发给家长

## 🗂 项目结构

```
├── class-report/        # 报告归纳规则原始 Skill（LLM 提示词的来源）
├── webapp/              # 应用主体
│   ├── server.py           # 本地后端（标准库 http.server，零框架）
│   ├── llm_client.py       # LLM 调用：报告 JSON / 脑图 Mermaid / 家长反馈 / 学情总结
│   ├── report_builder.py   # python-docx 排版（品牌、水印、脑图 SVG 嵌入）
│   ├── mermaid_render.py   # Mermaid → Edge 无头渲染 → 文字化 SVG → 智能切片 PNG
│   ├── pdf_convert.py      # docx → pdf（Word COM 优先，LibreOffice 兜底）
│   ├── static/index.html   # 全部前端（玻璃拟态 UI，原生 JS + Tailwind）
│   ├── config.example.json # 配置模板（真实 config.json 不上传）
│   ├── 启动/停止课堂反馈生成器.vbs
│   └── HANDOVER.md         # 开发交接文档（完整实现细节与踩坑记录）
└── 总结.md              # 整体工作总结
```

## 🔒 隐私与安全

- `config.json`（API Key）、`brand/`（Logo）、`runtime/`（报告数据，含学生姓名）均在 `.gitignore` 中，仅存本机
- 除调用你配置的 LLM 端点外，不访问任何外部服务

## ⚠️ 已知说明

- Mermaid 脑图过长时按知识分支间隙切成多段分页，跨段连线在切片边缘截断属正常现象
- 推理型模型（如 deepseek-v4-pro）的思考过程占用 max_tokens，本项目各调用已预留余量

## License

MIT
