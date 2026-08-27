# -*- coding: utf-8 -*-
"""LLM 调用（OpenAI 兼容协议）：把「课程速览 + 课程纪要」归纳成课堂报告结构化 JSON。

支持 DeepSeek / Kimi(Moonshot) / 阿里云百炼(兼容模式) / 任意 OpenAI 兼容端点，
差异仅在 base_url + api_key + model，协议完全一致。
归纳规则全部来自 class-report/SKILL.md（九大板块、emoji 标题、剔除学科无关内容等）。
"""
import json
import urllib.request

SYSTEM_PROMPT = """你是一位专业的课堂报告撰写助手。根据老师提供的「课程速览」（带时间戳的课堂分段小结）与「课程纪要」（课堂逐字稿/纪要摘要），归纳整理出一份结构化课堂报告内容，以严格 JSON 输出。

【硬性规则】
1. 封面五要素：标题、副标题、学科、授课形式、主讲内容、报告日期。不得出现「学生升学目标」等个性化字段。
2. 全文不得出现「由 AI 生成 / AI 辅助生成」等字样。
3. 「课堂内容」要精简：只写核心结论概览（4~6 条），不铺陈细节。
4. 「课程纪要」须大幅度保留原文细节（关键数据、实验、推导、例题），按知识点分小节，仅删除与学科学习无关的内容（升学目标/竞争焦虑、AI 工具闲聊、老师个人趣闻、其它学科话题等一律剔除）。
5. 「课堂作业」从纪要「作业布置」中提取作业范围与要求（如书名+页码），并附答题规范提醒；不编造、不夸大作业量。
6. 「课堂表现」分「优点」「待改进」两段，客观具体（各 3 条左右），围绕专注度、概念辨析、答题规范等，不空泛、不过度夸赞。
7. 「易错分析」每条给出「误区标签 + 正确理解/纠正」，突出易错字、易混概念、易漏步骤。
8. 「学习建议」可操作、落地，只写学科学习方法，不写放松心态/心理建设/家长沟通等内容。
9. 各小节标题需带一个贴切的 emoji（课程纪要小节标题形如「🔭 1. 声音的产生」）。

【输出 JSON Schema（严格遵守，不要输出任何额外文字）】
{
  "meta": {
    "title": "课程主标题（如：声音是什么）",
    "subtitle": "副标题，形如「—— 章节名（第X章）」",
    "subject": "学科（年级 · 章节），如：初中物理（八年级 · 声现象）",
    "form": "授课形式，如：一对一辅导课",
    "content": "主讲内容一句话概括",
    "date": "报告日期，形如 2026 年 8 月 22 日"
  },
  "mainline": "一句话主线说明（本节围绕「…」展开，核心结论如下：）",
  "content": [{"lead": "结论标签：", "text": "核心结论内容"}],
  "performance": {
    "pros": [{"lead": "标签：", "text": "具体表现"}],
    "cons": [{"lead": "标签：", "text": "具体表现"}]
  },
  "minutes": [{"title": "emoji + 序号 + 小节标题", "points": ["逐条保留的纪要细节"]}],
  "gains": [{"lead": "标签：", "text": "收获说明"}],
  "homework": [{"lead": "作业：", "text": "作业范围与要求"}, {"lead": "规范：", "text": "答题规范提醒"}],
  "mistakes": [{"lead": "误区标签：", "text": "正确理解/纠正"}],
  "suggestions": [{"lead": "建议：", "text": "具体可落地的学习方法"}]
}

【数量要求】content 4~6 条；pros/cons 各 3 条左右；minutes 按知识点 4~8 个小节；gains 6~8 条；homework 2~4 条；mistakes 5~8 条；suggestions 5~7 条。
若用户提供了封面字段（课程标题/学科/授课形式/报告日期），优先使用用户提供的值。"""


def generate_report_json(base_url, api_key, model, overview, notes,
                         meta_hints=None, extra_body=None):
    user_msg = "【课程速览】\n%s\n\n【课程纪要】\n%s" % (overview, notes)
    if meta_hints:
        hints = "、".join("%s=%s" % (k, v) for k, v in meta_hints.items() if v)
        if hints:
            user_msg = "【用户指定的封面字段】%s\n\n%s" % (hints, user_msg)

    url = base_url.rstrip("/") + "/chat/completions"
    req_body = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
        "response_format": {"type": "json_object"},
        "max_tokens": 8000,
        "temperature": 0.3,
    }
    if extra_body and isinstance(extra_body, dict):
        req_body.update(extra_body)  # 供应商特定参数，如 {"thinking": {"type": "disabled"}}
    body = json.dumps(req_body).encode("utf-8")

    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/json",
        "Authorization": "Bearer " + api_key,
    })
    try:
        with urllib.request.urlopen(req, timeout=240) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "ignore")[:500]
        raise RuntimeError("模型接口返回 HTTP %s：%s" % (e.code, detail))
    text = payload["choices"][0]["message"]["content"]
    return json.loads(text)


REQUIRED_KEYS = ["meta", "content", "performance", "minutes", "gains",
                 "homework", "mistakes", "suggestions"]


# ---------------- 知识点脑图（独立第二次调用） ----------------
MINDMAP_PROMPT = """你是课堂知识点脑图绘制助手。根据老师已归纳好的课堂报告 JSON，绘制一份知识点脑图，输出一段 Mermaid 代码。

【输出格式】
- 只输出 Mermaid 代码本身：不要任何解释文字，不要用 ``` 包裹，不要输出 JSON。

【版式（竖版长图）】
- 第一行固定为：graph LR
- 层级从左向右展开，分支纵向排开，整体呈竖版长图（根在左，枝叶向右、向下铺）。
- 所有节点一律使用圆角矩形语法：节点id("节点文本")，文本用中文。

【代码组织方式（严格遵守，参照文末示例）】
- 开头是「固定样式配置区」：六行 classDef 原样照抄（见示例）。
- 节点 id 用语义化层级命名：根节点固定叫 Root；一级分支依次叫 A、B、C、D……；某分支的二级节点叫 A1、A2……；三级叫 A1_1、A1_2……；四级叫 A1_1_1……（用父 id 做前缀，一眼看出层级归属）。
- 每个节点定义时就地绑定样式，用行内语法：`A1("大题过程书写至关重要"):::level2` —— 禁止使用单独的 class 语句。
- 按分支分组书写：先把 Root 和一级分支写完，然后逐个分支处理——每个分支一组，组内先定义该分支的节点、再写该分支的连线；组与组之间空一行，可用 %% 注释标明分支名。
- 层级与样式对应：Root → :::rootLevel，一级 → :::level1，二级 → :::level2，三级 → :::level3，四级 → :::level4，五级 → :::level5。

【层级深度（重要：按内容自然展开，不要死板对齐）】
- 层级深度由内容决定，2 级到 5 级都可以：细节多的分支往深处展开（可到四级、五级），内容少的分支到二级就结束。
- 不同分支的深度应当不同，禁止把每个分支都凑成同样的三层结构。
- 禁止为了凑层级而拆出无信息量的中间节点；也禁止把多个独立知识点挤进一个节点文本。

【内容要求（最重要）】
- 根节点 = 报告主标题（meta.title）。
- 一级分支必须完整覆盖课程纪要（minutes）中的每一个知识点小节，课堂内容（content）的结论也要纳入对应分支，一个小节也不许漏。
- 深层节点展开细节：定义、公式、实验现象、例题结论、易错点、作业要求等，尽量保留报告中的具体信息（数字、公式、实例），不要只写空泛标题。
- 节点总数参考 25~60 个，内容多就多写，不设硬上限；宁可多挂一层，也不要漏掉知识点。
- 节点文本精简（≤ 20 字）。文本内禁止英文双引号和英文括号；冒号、引号等标点一律用全角（：「」（）），避免破坏 Mermaid 语法。

【连线要求（缺一不可，否则脑图没有结构）】
- 必须用 `-->` 把所有节点连成一棵树：除 Root 外，每个节点都必须有且只有一个父节点指向它。
- 禁止只列节点不连线。节点数为 N 时，连线数必须是 N-1。

【完整示例（代码组织方式以此为准，内容仅示意；注意各分支深度不同）】
graph LR
    %% ----- 固定样式配置区 -----
    classDef rootLevel fill:#1e3a8a, stroke:#1e3a8a, stroke-width:2px, color:#ffffff;
    classDef level1 fill:#e0f2fe, stroke:#0284c7, stroke-width:2px, color:#075985;
    classDef level2 fill:#fef3c7, stroke:#d97706, stroke-width:1.5px, color:#92400e;
    classDef level3 fill:#f3f4f6, stroke:#9ca3af, stroke-width:1px, color:#374151;
    classDef level4 fill:#ffffff, stroke:#d1d5db, stroke-width:1px, color:#4b5563;
    classDef level5 fill:#fafafa, stroke:#e5e7eb, stroke-width:1px, color:#6b7280;

    %% ----- 根节点与一级分支 -----
    Root("物理学习核心要点"):::rootLevel
    A("声现象复习"):::level1
    Root --> A
    B("光现象入门"):::level1
    Root --> B

    %% ----- 分支 A（内容多，展开到四级） -----
    A1("超声波"):::level2
    A2("次声波"):::level2
    A --> A1
    A --> A2
    A1 --> A1_1("特点：方向性好、穿透力强"):::level3
    A1 --> A1_2("应用"):::level3
    A1_2 --> A1_2_1("声呐、B超"):::level4
    A1_2 --> A1_2_2("清洗、碎石"):::level4
    A2 --> A2_1("定义：低于20Hz"):::level3

    %% ----- 分支 B（内容少，到二级即止） -----
    B1("光源：本身发光的物体"):::level2
    B2("光的三原色：红绿蓝"):::level2
    B --> B1
    B --> B2"""


def _strip_code_fences(text):
    text = text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        # 去掉首行 ```mermaid 和尾行 ```
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines).strip()
    return text


def _check_mindmap_code(code):
    """结构校验：必须 graph/flowchart 开头，且连线数 ≥ 节点数-1（否则渲染成无结构的单列）。"""
    if not (code.startswith("flowchart") or code.startswith("graph")):
        raise ValueError("模型未返回合法 Mermaid 代码: %r" % code[:80])
    nodes = code.count('("')
    edges = code.count("-->")
    if nodes >= 2 and edges < nodes - 1:
        raise ValueError("脑图缺连线（节点 %d、连线 %d）" % (nodes, edges))


def generate_mindmap_mermaid(base_url, api_key, model, data, extra_body=None):
    """第二次独立调用：基于已归纳的报告 JSON 生成完整知识点脑图的 Mermaid 代码。
    结构校验不合格时带反馈重试一次。"""
    payload = {k: v for k, v in data.items() if k != "mindmap_mermaid"}
    user_msg = "【课堂报告 JSON】\n" + json.dumps(payload, ensure_ascii=False)

    url = base_url.rstrip("/") + "/chat/completions"
    messages = [
        {"role": "system", "content": MINDMAP_PROMPT},
        {"role": "user", "content": user_msg},
    ]
    last_err = None
    for attempt in range(2):
        req_body = {
            "model": model,
            "messages": messages,
            # 推理型模型（如 deepseek-v4-pro）的思考过程也占 max_tokens，
            # 4000 容易被 reasoning 耗尽导致 content 为空，给足余量
            "max_tokens": 16000,
            "temperature": 0.2,
        }
        if extra_body and isinstance(extra_body, dict):
            req_body.update(extra_body)
        body = json.dumps(req_body).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + api_key,
        })
        try:
            with urllib.request.urlopen(req, timeout=240) as resp:
                resp_payload = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "ignore")[:500]
            raise RuntimeError("脑图生成接口返回 HTTP %s：%s" % (e.code, detail))
        msg = resp_payload["choices"][0]["message"]
        code = _strip_code_fences(msg.get("content") or "")
        if not code.strip():
            fr = resp_payload["choices"][0].get("finish_reason")
            raise RuntimeError(
                "脑图生成返回空内容（finish_reason=%s）——推理型模型的思考过程会占 max_tokens，"
                "若是 length 说明额度被推理耗尽" % fr)
        try:
            _check_mindmap_code(code)
            return code
        except ValueError as e:
            last_err = e
            if attempt == 0:
                messages.append({"role": "assistant", "content": code})
                messages.append({"role": "user", "content":
                    "你上次输出的 Mermaid 代码缺少 --> 连线，所有节点是孤立的，渲染出来没有结构。"
                    "请重新输出完整代码：除根节点外每个节点都必须被一个父节点用 --> 指向，"
                    "节点数为 N 时连线数必须是 N-1。只输出 Mermaid 代码。"})
    raise last_err


def validate(data):
    missing = [k for k in REQUIRED_KEYS if k not in data]
    if missing:
        raise ValueError("模型返回缺少字段: %s" % ", ".join(missing))
    for k in ["title", "subtitle", "subject", "form", "content", "date"]:
        if k not in data["meta"]:
            raise ValueError("模型返回 meta 缺少字段: %s" % k)
    if "pros" not in data["performance"] or "cons" not in data["performance"]:
        raise ValueError("模型返回 performance 缺少 pros/cons")
    # 脑图/家长反馈为可选字段：非字符串则丢弃，不影响主报告
    for opt in ("mindmap_mermaid", "parent_feedback"):
        if opt in data and not isinstance(data[opt], str):
            data.pop(opt)
    return data


# ---------------- 学情总结（跨报告 AI 归纳） ----------------
STUDENT_SUMMARY_PROMPT = """你是一位专业的学情分析助手。根据某位学生历次课堂报告的摘要（按时间正序），输出一份跨报告的阶段性学情总结。

【分析依据（最重要的原则）】
- 学情判断必须以「课堂表现」（优点 / 待改进）为核心依据：学生的专注度、互动、概念掌握、答题规范等课堂实录表现。
- 不得以错题、误区、易错点作为评价学生的依据；作业与收获仅作辅助参考。

【硬性规则】
1. 只基于给定的报告内容归纳，不编造学生表现；报告数量少、数据不足时如实说明趋势尚不明显。
2. 全文不得出现「由 AI 生成 / AI 辅助生成」等字样。
3. 输出纯文本：不要 JSON、不要 Markdown 语法（不用 # * 等符号）。分四个小节，小节标题原样使用（独占一行）：
📈 进步轨迹
🌟 课堂亮点
⚠️ 待改进方向
💡 教学建议
4. 每小节 2~6 条，每条独占一行并以「・」开头；内容要具体，引用报告中课堂表现的真实细节，不空泛。
5. 「进步轨迹」按时间顺序对比各节课课堂表现的变化；「课堂亮点」归纳反复出现的优点；「待改进方向」合并多次报告中重复出现的待改进项；「教学建议」给出可落地的后续教学安排。"""


def generate_student_summary(base_url, api_key, model, name, digest, extra_body=None):
    """基于某学生历次报告摘要（list[dict]，时间正序）生成阶段性学情总结纯文本。"""
    user_msg = "【学生姓名】%s\n\n【历次课堂报告摘要（按时间正序）】\n%s" % (
        name, json.dumps(digest, ensure_ascii=False))

    url = base_url.rstrip("/") + "/chat/completions"
    req_body = {
        "model": model,
        "messages": [
            {"role": "system", "content": STUDENT_SUMMARY_PROMPT},
            {"role": "user", "content": user_msg},
        ],
        "max_tokens": 3000,
        "temperature": 0.3,
    }
    if extra_body and isinstance(extra_body, dict):
        req_body.update(extra_body)
    body = json.dumps(req_body).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/json",
        "Authorization": "Bearer " + api_key,
    })
    try:
        with urllib.request.urlopen(req, timeout=240) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "ignore")[:500]
        raise RuntimeError("学情总结接口返回 HTTP %s：%s" % (e.code, detail))
    return payload["choices"][0]["message"]["content"].strip()


# ---------------- 家长反馈（简易文字版，直接发微信） ----------------
PARENT_FEEDBACK_PROMPT = """你是培训机构的授课老师，需要根据课堂报告 JSON，写一段直接发给家长的简易文字反馈（微信消息风格）。

【输出格式（严格遵守：纯文本，不要 JSON、不要 Markdown、不要代码块、不要任何解释）】
{学生姓名}家长您好，我是孩子的{科目名}老师，以下是孩子近期{科目名}课堂学习情况反馈，请查收：

【授课年级及科目】
{年级}{科目名}
【课堂授课内容】
1. ……
2. ……

【课堂整体表现⭐️】
……

【学习建议】
……

【课后作业】
……

【内容要求】
- 学生姓名、年级、科目名取自 meta（meta.student / meta.form / meta.subject 中的年级与学科信息；subject 形如「初中物理（八年级 · 声现象）」时，科目名取「物理」、年级取「八年级」）。
- 【课堂授课内容】按课程纪要（minutes）的小节逐条编号列出，每条「主题：要点概括」，覆盖所有小节，不遗漏。
- 【课堂整体表现⭐️】以课堂表现（performance）的 pros/cons 为核心依据，结合课堂内容（content）结论：先总述上课状态，再按知识板块分段写掌握情况；语气亲切、以鼓励为主，不足之处理性委婉表达（如「仅少数偏难题型掌握不够熟练，后续需要针对性加强练习」），不夸大、不点名批评。
- 【学习建议】依据学习建议（suggestions）改写为家长易懂的表达，可操作、落地。
- 【课后作业】依据课后作业（homework）逐条列出；若无作业信息则写「完成对应课后习题」。
- 全文不得出现「由 AI 生成 / AI 辅助生成」等字样；不得出现 JSON、字段名、「纪要」「速览」等内部术语。
- 篇幅控制在 200~400 字（不含格式标题）。"""


def _check_parent_feedback(text):
    """结构校验：五个固定板块标题齐全，且含称呼行。"""
    for tag in ("家长您好", "【授课年级及科目】", "【课堂授课内容】",
                "【课堂整体表现⭐️】", "【学习建议】", "【课后作业】"):
        if tag not in text:
            raise ValueError("家长反馈缺少板块: %s" % tag)


def generate_parent_feedback(base_url, api_key, model, data, extra_body=None):
    """第三次独立调用：基于已归纳的报告 JSON 生成家长版简易文字反馈。
    结构校验不合格时带反馈重试一次。"""
    payload = {k: v for k, v in data.items()
               if k not in ("mindmap_mermaid", "parent_feedback")}
    user_msg = "【课堂报告 JSON】\n" + json.dumps(payload, ensure_ascii=False)

    url = base_url.rstrip("/") + "/chat/completions"
    messages = [
        {"role": "system", "content": PARENT_FEEDBACK_PROMPT},
        {"role": "user", "content": user_msg},
    ]
    last_err = None
    for attempt in range(2):
        req_body = {
            "model": model,
            "messages": messages,
            # 推理型模型（如 deepseek-v4-pro）的思考过程也占 max_tokens，给足余量
            "max_tokens": 8000,
            "temperature": 0.4,
        }
        if extra_body and isinstance(extra_body, dict):
            req_body.update(extra_body)
        body = json.dumps(req_body).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + api_key,
        })
        try:
            with urllib.request.urlopen(req, timeout=240) as resp:
                resp_payload = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "ignore")[:500]
            raise RuntimeError("家长反馈接口返回 HTTP %s：%s" % (e.code, detail))
        msg = resp_payload["choices"][0]["message"]
        text = _strip_code_fences(msg.get("content") or "").strip()
        if not text:
            fr = resp_payload["choices"][0].get("finish_reason")
            raise RuntimeError(
                "家长反馈返回空内容（finish_reason=%s）——推理型模型的思考过程会占 max_tokens" % fr)
        try:
            _check_parent_feedback(text)
            return text
        except ValueError as e:
            last_err = e
            if attempt == 0:
                messages.append({"role": "assistant", "content": text})
                messages.append({"role": "user", "content":
                    "你上次的输出缺少固定板块标题。请严格按格式重新输出：称呼行 + "
                    "【授课年级及科目】【课堂授课内容】【课堂整体表现⭐️】【学习建议】【课后作业】五个板块，"
                    "只输出纯文本。"})
    raise last_err
