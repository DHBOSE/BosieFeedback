# 课堂反馈生成器 · 工作交接文档

> 最后更新：2026-08-26 ｜ 状态：全部功能可用，端到端测试通过（学生姓名必填、学情分析以课堂表现为依据、支持课程范围选择、设置弹窗遮挡修复、删除顶部型号字样）

## 项目简介

本地 Web 应用：老师粘贴「课堂速览 + 课堂纪要」（可附脑图图片）→ AI 按 `class-report` SKILL 规范归纳 → 生成结构化课堂报告 → 网页内预览真实 Word 排版 → 导出 Word / PDF。

灵感与规范来源（工作区根目录）：
- `class-report/SKILL.md` + `templates/`：报告生成规范与 Word 模板（本项目 `report_builder.py` 完整复刻其版式）
- `stitch_ (1)/`：Liquid Glass 设计稿（前端视觉来源）
- 根目录的「初中物理-声现象-声音是什么-课堂报告-修订版.docx/pdf」是参照样例

## 目录结构（E:\BosieFeedback\webapp\）

| 文件 | 作用 |
|---|---|
| `server.py` | 后端（Python 标准库 http.server，零第三方依赖）。端口 7100 |
| `llm_client.py` | LLM 调用（OpenAI 兼容协议）+ SKILL 规则 system prompt + JSON 校验（含 mindmap_mermaid 脑图字段） |
| `mermaid_render.py` | Mermaid → PNG 离线渲染（本机 Edge 无头：dump-dom 提取 SVG → 3x 截图；语法错误/无 Edge 自动兜底） |
| `report_builder.py` | python-docx 生成 Word（封面无框线表格、emoji 标题、八大板块、脑图嵌入） |
| `pdf_convert.py` | Word COM 转 PDF（LibreOffice 兜底；本机 LibreOffice 报 0xC0000142 不可用） |
| `config.json` | 多供应商配置（仅存本地，接口只返回掩码 key）+ `brand` 段（水印：org_name/contact + cover/header/footer 三槽位 enabled/height_cm） |
| `brand/` | 水印图片：cover.png|jpg / header.png|jpg / footer.png|jpg（由水印窗口上传管理；旧版 logo.* 启动时自动迁移为 cover+header） |
| `static/index.html` | 前端单页（全部逻辑内联） |
| `static/vendor/` | 本地化前端依赖：tailwind.js、jszip.min.js、docx-preview.min.js |
| `runtime/<job>/` | 每次生成一个目录：data.json / report.docx / report.pdf / mindmap_slices.json + mindmap_N.png/.svg（脑图切片；旧版为单图 mindmap.png）/ filename.txt |
| `runtime/_students/` | 学情功能产物：AI 学情总结缓存 <姓名>.json、学情趋势 docx/pdf（`_` 开头目录不参与报告扫描） |
| `test_e2e.py` | 端到端回归测试（8 步） |
| `test_brand.py` | 机构品牌回归测试（保存/取图/docx 断言/清理） |
| `test_batch.py` | 批量打包回归测试（zip 内容/重名序号/非法 job 校验） |
| `启动课堂反馈生成器.vbs` | **双击启动**（无黑窗口、自动开浏览器、已在运行则不重复启动） |
| `停止课堂反馈生成器.vbs` + `stop_server.ps1` | 双击停止（只匹配本项目 server.py） |
| `package.json` | `npm run dev` 供 Kimi Work 预览卡片使用 |

## 启动方式

- 日常使用：双击 `启动课堂反馈生成器.vbs`
- Kimi Work 预览：`npm run dev` → http://localhost:7100/
- 停止：双击 `停止课堂反馈生成器.vbs`

## 已实现功能

1. **生成**：速览+纪要（必填）+ 脑图（点击/拖拽/Ctrl+V 粘贴）+ 封面四字段（可选，留空 AI 归纳）
2. **预览**：docx-preview 渲染真实 Word 文件，CSS `zoom` 自动适配栏宽（清晰、无裁剪），窗口变化自动重适配
3. **编辑后重新排版**：左栏结构化编辑器（全板块可改、可增删条目）→「应用修改并刷新预览」→ `POST /api/update` 重建 docx（旧 PDF 自动作废）
4. **导出**：下载 Word / 导出 PDF（Word COM，文件名自动按「学科-标题-课堂报告」）
5. **多模型供应商**：DeepSeek（默认 deepseek-v4-pro）/ Kimi / 百炼 / 自定义 OpenAI 兼容；生成按钮上方快切下拉 + 右上角 ⚙ 设置面板（改 key/模型/默认供应商，保存即生效）
6. **历史报告**：右上角「历史报告」弹窗：列出 runtime/ 下所有报告，支持载入继续编辑、下载 Word/PDF、删除（带确认）
7. **水印设置**：顶栏「水印」按钮 → 独立窗口。机构名 + 联系方式；封面/页眉/页脚三槽位各自独立上传图片（PNG/JPG，不限大小）、各自调图片高度（0.3–4cm 滑杆）、各自启用开关。封面：图片+机构名横排居中；页眉：机构名+图片靠右（细分隔线）；页脚：机构名 · 联系方式 · 第 X 页 · 图片全部靠右。**文档分封面/正文两节：封面节无页眉页脚，正文节页码从 1 重新开始（pgNumType start=1）**。「刷新预览」用未保存的当前设置生成真实 docx 并在窗口右侧 docx-preview 渲染（不消耗 LLM、不落盘）。品牌全局生效，旧报告重新导出时用最新设置。**正文大水印（2026-08-26）**：水印窗口第四个卡片「正文水印」——启用开关 + 文字/图片两种类型 + 大小（占版心宽 10–150%）/不透明度（3–100%，越小越浅）/倾斜角度（-90°~90°，默认 -45°）三滑杆；文字型留空则用机构名，图片型走第四个隐藏槽位 `body`（brand/body.png|jpg）。实现：服务端 PIL 把 文字（300px 微软雅黑渲染→裁剪）或图片 → 旋转 → 调 alpha **烘焙成一张透明底 PNG**，`report_builder._add_body_watermark(sec2, brand)` 在正文节页眉插入 `wp:anchor behindDoc="1"` 浮动图（`_float_behind_run`，相对页面水平/垂直居中、wrapNone）——封面页无水印，正文每页居中。docx-preview 能正常渲染该 anchor，预览可见；PDF 转换实测水印只出现在正文页且精确居中。两个 build_* 都调用（页眉未启用时也会 unlink 页眉单独放水印）。**注意回退链：文字型水印文本留空 → 用机构名 → 机构名也为空则无内容可渲染，docx 里静默没有水印**——前端 previewBrand 已加橙色提示「文字水印没有内容…」（2026-08-26，用户曾因此误以为功能失效）。**水印字体（2026-08-27）**：文字型新增字体下拉（玻璃样式，选项按各自字体渲染预览字形）——`report_builder.WM_FONT_TABLE` 23 种常见中文字体（微软雅黑/宋体/新宋体/黑体/仿宋/楷体/等线/隶书/幼圆 + Office 华文全家桶 + 方正舒体/姚体），`(key, 显示名, 文件名, ttc index, CSS font-family)` 五元组，`available_wm_fonts()` 按文件存在性过滤（新宋体在 simsun.ttc 的 index 1）；配置存 `watermark.font`（默认 msyh，非法 key 回退 msyh），`GET /api/brand` 返回 `wm_fonts` 清单；`_wm_font_file` 解析失败时回退第一个可用字体。**水印颜色（2026-08-27）**：`watermark.color` 存 6 位 hex（默认主题青 216873，非法值回退），仅文字型生效；前端为原生取色器 + 5 个预设色圆点（216873/1e3a8a/E65A5A/787878/333）+ hex 标签
8. **脑图等比缩放**：嵌入脑图宽度优先顶满 15.5cm，高度封顶 21.5cm，全程等比不变形
9. **批量生成 + zip 打包**：模型快切旁「批量生成」→ 弹窗内添加多节课（各自速览+纪要，**每课可单独附脑图**）→ 用当前模型逐课顺序生成（每课卡片实时显示 等待/生成中/✓文件名/✗错误，未填的自动跳过，单课失败不阻塞）→「打包下载 ZIP」（勾选可连 PDF 一起打，PDF 懒转换且单个失败不阻塞）。zip 内文件名即报告文件名，重名自动加 -2 序号
10. **AI 自动生成脑图**：未提供脑图时，报告 JSON 生成后做**第二次独立 LLM 调用**（`llm_client.generate_mindmap_mermaid`）：以归纳好的报告 JSON 为上下文，专心生图——`graph LR` 竖版长图（层级向右、分支纵向铺开）、圆角矩形、minutes 每个小节全覆盖、细节到定义/公式/实验/例题/作业。Mermaid 代码必须按用户给的模板风格书写（MINDMAP_PROMPT 内含完整示例）：开头四行固定 classDef（rootLevel/level1/level2/level3 配色）；**节点 id 语义化层级命名**（Root、A/B/C…一级、A1/A2 二级、A1_1 三级）；**行内绑定样式** `A1("文本"):::level2`，禁止单独 class 语句；按分支分组书写并用 `%%` 注释分隔；连线必须成树（N 节点 N-1 条 `-->`）；节点文本 ≤20 字、禁英文双引号/括号、标点全角。服务端 `mermaid_render.py` 用 Edge 无头渲染 PNG（3x）嵌入；生成与重排都走这套；渲染失败兜底空白占位。用户已传脑图则跳过第二次调用。批量模式同样生效。注意：模型常超出提示词给的节点数建议（实测 84 节点），出图偏长但结构正确（用户明确不加节点数硬上限 2026-08-24）。**渲染后按一级分支间隙智能切片分页**（见踩坑 16）：长图不再缩成一页窄条，而是切成若干段、每段顶满 15.5cm 版心宽各占一页，等效字号从 ~2.4pt 提到 ~8.4pt；Mermaid 布局调紧凑（nodeSpacing:18, rankSpacing:32）。**层级自适应 + 溢出修复（2026-08-27）**：提示词新增【层级深度】规则块——2~5 级按内容自然展开、各分支深度必须不同、禁止凑三层、禁止无信息量中间节点；classDef 扩到六级（新增 level4 白底 #d1d5db、level5 #fafafa/#e5e7eb），节点数改「参考 25~60，不设硬上限」；server 端 `ensure_mindmap` 补齐逻辑从「没有 rootLevel 就全量补」改为**按 classDef 名字逐条补**（防模型只写 level1-3 却用 level4）。渲染端三处修复：①`_svg_textify` 长文本按框宽**多行重排**（`_wrap_label` 像素贪心换行，行距 24px 与 line-height:1.5 一致）——根治长标签横向溢出框线；②text 颜色改用内联 `style="fill:..."`（CSS 优先级高于 fill 属性，否则被 mermaid 样式表 `#mm .label text{fill:#333}` 压成深灰，根节点深底深字不可见）；③`_node_rows` 切片高度取组内 rect 实际高度（多行节点 48/72px），`_slice_cuts` 无间隙窗口新增「穿过节点数最少」兜底扫描
11. **PDF 内嵌预览切换 + 直接打印**：预览栏「Word | PDF」分段切换。PDF 模式走 `/api/file?fmt=pdf`（无 download 参数即 inline，浏览器内置查看器 iframe 渲染）；前端 `fetchPdfUrl()` 先 fetch 校验 Content-Type 再建 blob URL（这样转换失败能拿到错误信息，iframe 直接加载做不到），同一 job 内复用不重复转换；「导出 PDF」与 PDF 预览共用这份 blob。「直接打印」：屏外定位 iframe（**不能 display:none，PDF 查看器不渲染**）加载 blob 后 `contentWindow.print()`，20s 超时兜底新标签页打开。任何重排/新生成（`refreshPreview`）都会 `invalidatePdf()` 并切回 Word 视图——服务端 /api/update 会删旧 pdf，不重置会展示过期内容。PDF 加载时预览区显示 **cell ripple 加载动画**（9 格绿色方块，`setPreviewLoading(text, true)` 第二参数切换 spinner/cell-loader）
12. **家长反馈 · 简易版（2026-08-27）**：报告生成后做**第三次独立 LLM 调用**（`llm_client.generate_parent_feedback`，max_tokens 8000 防推理模型耗尽）——按用户给的固定模板输出纯文本：称呼行「{学生}家长您好，我是孩子的{科目}老师…请查收：」+【授课年级及科目】【课堂授课内容】（minutes 小节逐条编号全覆盖）【课堂整体表现⭐️】（以 performance 为核心依据，先总述再按板块分段，不足委婉表达）【学习建议】【课后作业】；`_check_parent_feedback` 校验六要素，缺板块带反馈重试一次。结果存 `data["parent_feedback"]`（validate 视为可选字符串字段），失败静默不阻断主报告。前端：结果区网格内第三张全宽卡片（`xl:col-span-2`），可编辑 textarea +「一键复制」（clipboard API + execCommand 回退，按钮变「已复制 ✓」1.5s）+「AI 重新生成」（走 `POST /api/parent-feedback` {job,provider?,model?}，旧报告可补生成，写回 data.json）；`collectEditor` 带上 parent_feedback，点「应用修改」时一并保存。生成与 openReport 载入都会填充 textarea。**批量流程同步支持**：批量走前端循环调 /api/generate，每份报告自动产出家长反馈；条目成功后卡片底部出现可折叠「家长反馈」块（展开可编辑、单独一键复制）；zip 打包时每份报告附 `<报告名>-家长反馈.txt`（server 端 handle_batch_zip 从 data.json 读取；若前端传了 `feedbacks{job:text}` 则以弹窗里修改后的版本为准），单个附件失败不阻塞打包（2026-08-27）。**批量条目支持单独封面信息（2026-08-27）**：每节课卡片在学生姓名下方新增可折叠「封面信息（可选，不填则由 AI 归纳）」（en.meta{title,subject,form,date} + en.metaOpen），四字段任一非空才作为 reqBody.meta 传给 /api/generate；服务端 meta_hints 链路原生支持，无需改动
12. **UI 样式改版（2026-08-24）**：①全站输入框改为**浮动标签**样式：`.formField` 包裹 `input/textarea + span`，灰底 #f1f1f1、聚焦 2px 深色描边、聚焦/有值时标签**浮到输入框上方外部**（`translateY(calc(-100% - 2px)) scale(0.85)` + 去描边；原 snippet 的框内浮起 -12px 只适用于单行 input，textarea 内容多时会与首行文字重叠——2026-08-24 修复）。用 `placeholder=" "` + `:not(:placeholder-shown)` 判定有值，**JS 赋值也能触发浮动**。`.formField` 自带 margin-top:18px 给浮起标签留位。②液态玻璃按钮 `.lq-wrap > .lq-btn > .lq-label + .lq-shadow` 结构（`@property` 角度动画、conic 边框、`:has()` 阴影联动；`.lq-btn` 有 `all:unset`，加功能时注意）覆盖**全部主按钮**：生成/应用修改/下载 Word/导出 PDF/直接打印 + 弹窗里的保存设置/刷新预览/添加一节课/开始批量生成/打包 ZIP。弹窗输入框（设置 4 项、水印机构名/联系方式、批量条目 textarea）也全部改为 formField 浮动标签。保留旧样式未改的：左栏结构化编辑器的小输入框（自有「标签：」placeholder 模式）、历史弹窗每行小按钮、各弹窗关闭 X、文字链接、脑图上传/移除等小工具按钮——列表/紧凑场景 lq 大阴影不合适③`pdf_convert.py` 两个 subprocess 调用都加 `CREATE_NO_WINDOW`，消除 PDF 转换时的 PowerShell 黑框
13. **学生档案 + 学情趋势（2026-08-26）**：①**关联学生**：学生姓名**必填**（前后端双重校验；批量模式缺学生姓名的条目自动跳过），写入 `data.json` 的 `meta.student`；封面加「学生」行、文件名为「学科-标题-学生-课堂报告」（`server.report_filename`）；旧报告载入编辑器补填即归档。②**历史筛选**：历史弹窗顶部下拉按学生过滤。③**学生视图**：顶栏「学生」→ 学生卡片列表 → 学情趋势页：统计行（报告数/时间跨度/优点与待改进条数）+ **课程范围选择器**（第 N 节 · 日期 · 标题，默认全选，作用于 AI 总结与 PDF 导出，jobs 列表传给服务端过滤）+ 待改进汇总（带日期）+ 逐课时间线（第 N 节编号）。**误区/易错不作为学情分析依据**：趋势页不展示、digest 不含 mistakes、提示词明确以课堂表现（优点/待改进）为核心依据，四节为 进步轨迹/课堂亮点/待改进方向/教学建议。④**AI 学情总结**：`POST /api/student_summary`（可带 jobs 范围），缓存到 `runtime/_students/<姓名>.json`（含 jobs 范围）。⑤**导出学情 PDF**：`GET /api/student_pdf?name=&jobs=`，`build_student_report` 生成封面 + AI 总结（**仅当缓存 jobs 与本次范围一致才附**）+ 待改进汇总 + 逐课时间线。全程零数据库
14. **封面信息始终展开 + 玻璃下拉组件（2026-08-26）**：①封面信息不再用 `<details>` 收起，常驻展开（badge 图标 + 5 个 formField，学生姓名排第一）。②**原生 `<select>` 的选项弹层由 OS 渲染、完全不可定制样式**，故自绘玻璃下拉：`enhanceSelect(native, opts)` 隐藏原生 select 作数据源，生成 `.glass-select-trigger`（圆角胶囊）+ `.glass-select-panel`（position:fixed 挂 body、z-index 300、rgba(255,255,255,0.72) + backdrop-filter blur(24px)、圆角 14px、选项圆角 10px、选中高亮 #002fa7）；点击外部/页面滚动/resize 关闭，空间不足自动向上弹出。接入：模型快切 + 课程范围 from/to（「从…到…」同一行 flex-nowrap，其弹层同款玻璃效果）。**关键：任何 JS 重建原生 select 的 options 后必须调 `syncGlass(native)` 同步**（renderQuickSwitch、renderStudentDetail 已接）

## API 一览

- `GET  /api/providers` / `POST /api/config`：供应商读取/保存（key 掩码）
- `GET  /api/brand` / `POST /api/brand`：品牌读取/保存（body: org_name/contact + slots{cover,header,footer:{enabled,height_cm}} + **watermark{enabled,type:text|image,text,size_pct,opacity_pct,angle}** + images{slot:base64，slot 可为 body} + remove_images[slot 可为 body]；PNG/JPG 魔数校验）
- `GET  /api/brand/image?slot=cover|header|footer|body`：槽位图片（前端加 `?t=` 破缓存）
- `POST /api/brand/preview`：按请求体（未保存的）设置生成预览 docx 直接返回，不落盘配置/图片
- `POST /api/batch_zip`：{jobs: [...], include_pdf?} → zip 下载（中文文件名、重名加序号、PDF 懒转换）
- `POST /api/generate`：{overview, notes, student(必填), mindmap?, meta?, provider?, model?} → {job, filename, data}
- `POST /api/update`：{job, data} → 重建 docx
- `GET  /api/file?job=&fmt=docx|pdf[&download=1]`：文件下载/预览（PDF 懒转换）
- `GET  /api/history` / `GET /api/data?job=` / `POST /api/delete`：历史管理（history 响应含 student 字段）
- `GET  /api/students`：按学生聚合（姓名/报告数/最近日期），实时扫 runtime/
- `GET  /api/student?name=`：该学生全部报告（时间正序，完整 data）
- `POST /api/student_summary`：{name, provider?, model?, jobs?(课程范围)} → AI 学情总结（以课堂表现为依据；缓存 runtime/_students/<姓名>.json 含 jobs）
- `GET  /api/student_pdf?name=[&jobs=a,b,c]`：学情趋势 PDF 下载（每次重建；缓存总结的 jobs 与本次范围一致才附进 PDF）

## 关键决策与踩过的坑（重要！）

1. **DeepSeek 模型**：2026-04 起官方名为 `deepseek-v4-pro` / `deepseek-v4-flash`；旧名 `deepseek-chat`/`deepseek-reasoner` 是过渡路由，随时会关停。V4 默认开思考模式（慢），config.json 里 deepseek 配了 `extra_body: {"thinking": {"type": "disabled"}}` 保速度。
2. **预览模糊**：不要用 `transform: scale()`（位图缩放发虚），用 CSS `zoom`（矢量重排）。
3. **预览左裁切**：docx-preview 的 wrapper 是 flex 居中，页宽 > 容器宽会产生负偏移裁掉左边；已改 `display: block` + `margin: 0 auto`。
4. **pythonw 无窗口**：stdout 为 None，`server.py` main() 里已做 devnull 重定向，改动 main 时注意保留。
5. **VBS 必须 UTF-16 LE BOM 编码**（含中文时），ps1 保持纯 ASCII（PowerShell 5.1 把无 BOM UTF-8 当 ANSI 读）。
6. **Windows SO_REUSEADDR 允许重复绑端口**：Server 类已设 `allow_reuse_address = False`，绑定失败静默退出。
7. **托管 Python 的 PATH 没有 System32**：调 PowerShell 必须用完整路径（pdf_convert.py 里已处理）。
8. **前端依赖全部本地化**（vendor/），不要再引 CDN。
9. **URL 路由要先剥离查询串**再匹配（do_GET/do_POST 已处理）。
10. **封面/正文分两节（sectPr）**：封面节无页眉页脚部件；正文节 `is_linked_to_previous=False` 建自己的页眉页脚 + `pgNumType start=1` 让页码从正文起算。pgNumType 插入位置要在 w:cols 之前（schema 顺序）。
11. **Mermaid 渲染防语法破坏**：提示词里已禁止节点文本含英文双引号/括号/分号/冒号；渲染前用 `--dump-dom` 校验 `data-done="1"` 才截图，失败一律兜底留空，绝不让坏图进报告。Edge 路径写死在 mermaid_render.py 的 EDGE_CANDIDATES。
12. **模型会漏写 classDef 样式定义**（只绑定 class 不给定义 → 渲染成默认色）：server.py `ensure_mindmap` 渲染前检测并补齐四行 classDef（MINDMAP_CLASSDEFS），补齐后写回 data.json。注意：job 里已有 mindmap.png 时 update 不会重渲染，改样式需先删该文件。
13. **模型可能只列节点不画 `-->` 连线**（渲染成无结构竖条！）：提示词含「连线要求」+ 完整示例；`llm_client._check_mindmap_code` 接受 `graph` 或 `flowchart` 开头，校验 连线数 ≥ 节点数-1，不合格带反馈重试一次；`ensure_mindmap` 渲染前再校验一遍，不合格直接弃用留空。
14. **脑图矢量嵌入（解决 PDF 模糊，2026-08-24）**：两层修复——①渲染密度：旧逻辑 MAX_PX=4000 窗口上限会把长图 scale 压回 ~1x（如 4024px 高的图只出 144dpi 的 PNG）；现改用 `--force-device-scale-factor`（窗口保持 SVG 原生尺寸，密度最高 3x，单边像素上限 MAX_PIXELS=16000）。②矢量嵌入：渲染时同时存 `mindmap.svg`（`render_mermaid_to_png(..., out_svg=)`）；`report_builder._upgrade_picture_to_svg` 在 PNG 主 blip 下追加 `asvg:svgBlip` 扩展（uri `{96DAC541-...}`）指向 `word/media/mindmap-vector.svg` 部件——Word 转 PDF 按矢量渲染，docx-preview 和旧 Word 自动用 PNG 兜底，互不影响。**老报告升级方法：删掉 job 目录里的 mindmap_slices.json 和 mindmap_*.png/.svg（旧版单图则是 mindmap.png）再「应用修改并刷新预览」**，即按新逻辑重渲染
15. **mermaid SVG 的节点文字是 foreignObject（HTML），Word 不支持 → PDF 里文字全丢（只剩框和线）**！`htmlLabels:false` 也救不了（该版本节点标签仍用 foreignObject，只有边标签转原生）。根治：`mermaid_render._svg_textify` 在渲染后把每个 `<foreignObject>` 正则转成原生 `<text>`（提取 div style 里的颜色、单行短文本、YaHei 16px），PNG 与 SVG 走同一份转换结果。验证技巧：Office 渲染 SVG 用 `_svg2png.ps1`（PowerPoint COM 导入 SVG 导出 PNG 目检）；**别用 pypdf extract_text 验证文字完整性**——Word 把文字拆成多段 TJ run（CJK/拉丁分开），pypdf 按 ToUnicode 解码会假性「缺字」，数 TJ 数组里字形个数才准
16. **脑图智能切片分页（解决打印字太小，2026-08-24，①+② 组合方案）**：整图塞一页时高度封顶 21.5cm → 长图被缩成窄条，打印字约 2.4pt。①`mermaid_render.render_mermaid_slices(code, out_dir)`：`_node_rows` 解析各节点组 transform 的 y 区间，`_slice_cuts` 合并区间后找 >6px 的分支间隙做切点（贪心，每段高 ≤ w0×1.39 ≈ 21.5cm 等效），`_slice_svg` **同一份 SVG 只换 viewBox 裁剪 → 矢量无损**，产出 mindmap_N.png/.svg + mindmap_slices.json 清单。②布局调紧凑 nodeSpacing:18 / rankSpacing:32（全图宽 889→832px）。`report_builder` 脑图段落接受列表：多段时加说明行「脑图较大，已按知识分支分段排版…」+ 每段前分页符；`_upgrade_picture_to_svg` 部件名带序号 `mindmap-vector-%d.svg` 保证唯一。server `_existing_mindmaps(d)` 优先读切片清单、回退单图。实测 84 节点图切 3 片、每片顶满 15.5cm 宽，等效字号 16px/832CSSpx×15.5cm≈8.4pt。已知正常现象：跨分支的长连线在切片边缘被截断（viewBox 裁剪的必然结果）；切点只认节点组间隙，不保证语义完整分支但绝不把节点切两半
11. **test_brand.py 会清空品牌配置**：运行前先备份 config.json + brand/ 目录，测试 finally 里自动恢复（别中断进程，否则需手动从 runtime/_brand_backup 恢复）。
17. **学生姓名是精确匹配**：学情聚合按 `meta.student`（strip 后）字符串相等归档，「张三」和「张 三」会算两个学生；没有学生主数据表，改名只能靠编辑器逐个报告改。学情 PDF 每次实时重建（不做缓存），报告内容改了导出即最新
18. **设置弹窗 API Key 不能用动态 placeholder 显示掩码**：formField 浮动标签靠 `placeholder=" "` + `:not(:placeholder-shown)` 判定，一旦 JS 把 placeholder 改成「已保存：sk-…」，空值时标签浮不起来、与 placeholder 文字重叠遮挡。掩码提示必须放输入框外的独立 `<p>`（set-api-key-hint）。同理标签浮起需要足够上方空间：弹窗里用 `#provider-form .formField { margin-top: 30px }` 防止标签贴住上一行
19. **Windows 上 0.0.0.0:7100 与 127.0.0.1:7100 是两个独立 socket，可同时绑定**！用户的 VBS 启动器绑 127.0.0.1，测试时手动 `python server.py` 绑 0.0.0.0——localhost 请求永远走更精确的 127.0.0.1 那个（旧代码），造成「改了代码不生效」的假象。**测试前必须用 stop_server.ps1 杀掉 pythonw 实例**（它按 CommandLine 匹配 BosieFeedback，比按 PID 杀可靠），并用 `netstat -ano | grep :7100` 确认两种绑定都没了。前端静态文件按请求从磁盘读、后端逻辑在内存里——**不改代码重启服务时会出现「界面是新功能、导出是旧逻辑」的割裂状态**（2026-08-27 实例：用户选了仿宋水印但 PDF 仍是雅黑，原因就是服务没重启；配置其实已正确保存，重启后全链路验证通过）
20. **顶部不再显示当前型号**：nav-model 已按用户要求删除（2026-08-26）；当前模型只看生成按钮上方的快切下拉
21. **玻璃下拉的滚动关闭监听必须排除弹层内部**：`window.addEventListener("scroll", close, true)` 捕获阶段会收到弹层自身 overflow 滚动的 scroll 事件 → 选项一多、用户在弹层里一滚就关（2026-08-26 修复：`if (!panel.contains(e.target)) close()`）。另外用无头截图验证弹层时，`--virtual-time-budget` 期间页面自身的 scroll/resize 会把弹层关掉，截图看不到属预期——改用 `--dump-dom` 检查 `glass-select-panel` 无 hidden 类且有定位样式即可；CSS 渲染本身无问题（含 backdrop-filter 的同款面板在相同无头参数下截图正常）
22. **Bash 工具里 curl `-d` 直接写中文会变成 GBK 编码**（服务端 json 解析报 utf-8 decode 错误）：测试含中文的 POST 时，先用 python 把 body 写成 UTF-8 文件，再 `curl --data-binary @file`。同理无头 Edge `--screenshot` **不渲染 PDF**（空白页），验证 PDF 内容用 pypdf 解析 XObject/Do 调用；验证 docx 渲染效果用 file:// 协议 + `--allow-file-access-from-files` 加载 vendor docx-preview 截图
25. **`<textarea>` 用 JS 赋 `.value` 后，无头 Edge `--dump-dom` 序列化出来仍是空**（当前值是 DOM 属性，不进 innerHTML）——验证前端填充效果必须用 `--screenshot` 截图，别被 dump-dom 的空值误导（2026-08-27 实测）
24. **deepseek-v4-pro 是推理模型，reasoning 过程也占 max_tokens**：实测连「你好」都要 38 reasoning tokens；脑图生成 max_tokens=4000 会被推理耗尽导致 content 为空（finish_reason=length）。`generate_mindmap_mermaid` 已提到 16000，并对空 content 报明确错误（2026-08-27）。主报告生成若换推理模型也要检查 max_tokens
23. **mermaid 长文本节点靠 CSS `max-width:200px` 视觉换行，foreignObject 里没有 `<br/>` 标记**：textify 时必须按框宽重新换行。且**短标签的 foreignObject 宽度恰好等于文本宽度**（如「根」w=16，无 padding），统一 `w-10` 换行阈值会把 4 字短标签误拆两行——正确做法：估算宽 ≤ w+2 保持单行，只有触到 200px 上限的才按 w-4 换行。另外 `#mm .label text{fill:#333}` 这条 mermaid 自带 CSS 会覆盖 `<text fill>` 属性（任何 CSS 规则优先级都高于表现属性），文字颜色必须写 `style="fill:..."`（2026-08-27，三条均为实测根因）

## 调试钩子

- `http://localhost:7100/?autotest=1`：自动填样例并生成（端到端自测）
- `?autotest=history`：自动打开历史报告弹窗
- `?autotest=students`：自动打开学生档案弹窗（可加 `&name=<姓名>` 直接进该学生的趋势视图）
- `?autotest=load`：自动载入最新一条历史报告（可加 `&job=<id>` 指定报告）
- Office SVG 渲染目检：`_svg2png.ps1 -Svg <svg路径> -Png <输出.png> -PicW <宽> -PicH <高> -W <图宽> -H <图高>`（PowerPoint COM，验证 Word 同款 SVG 引擎的实际渲染效果）
- `?autotest=brand`：自动打开水印设置窗口并触发预览
- `?autotest=brandwm`：打开水印窗口、启用正文水印并预览，滚到第 2 页（验证正文水印链路）
- `?autotest=batch`：自动打开批量生成弹窗并填入两条样例
- `?autotest=pdfview`：载入最新历史报告后自动切到 PDF 预览（验证内嵌 PDF 链路）
- `?autotest=settings`：打开设置弹窗并填样例值（验证浮动标签浮起）
- `?autotest=select`：自动点开第一个玻璃下拉（验证弹层 DOM 与定位；无头截图里可能被虚拟时间的滚动关闭，以 `--dump-dom` 看 hidden 类为准）
- 无头截图验证：`msedge --headless --window-size=2850,1461 --virtual-time-budget=180000 --screenshot=out.png "http://127.0.0.1:7100/?autotest=1"`
- 回归测试：先确保 7100 空闲，`python test_e2e.py`（会污染 config.json 的 custom 段，测完需恢复）

## 待办候选

（学生档案 + 学情趋势已于 2026-08-26 交付，见功能 13；暂无新的候选，由用户提出后补充。）

> 「带意见重新生成」「脑图节点数硬上限」用户已明确不做（2026-08-24）。
