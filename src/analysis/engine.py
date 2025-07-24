import ollama
import json


def analyze_review_full(review_text: str) -> dict | None:
    """
    分析单条游戏长评，输出内容摘要、优点、缺点、标签。
    返回格式:
    {
        "summary": "内容摘要",
        "pros": "优点总结",
        "cons": "缺点总结",
        "tags": "标签1,标签2,标签3"
    }
    """
    truncated_text = review_text[:2000]

    prompt = f"""
    你是一位专业的游戏媒体编辑，擅长深入、全面地分析玩家长评。请仔细阅读下方评论，完成以下任务：

    **核心原则**：
    - **详尽分析**：尽可能提取和展现评论中的丰富信息，深入挖掘玩家的真实体验和观点
    - **实事求是**：只有在评论真正空洞无物（如纯表情符号、无意义重复）时才输出"该评论无有效内容"
    - **宁多勿少**：宁可输出详细丰富的内容，也不要过度精简而丢失信息
    - 仅输出中文、日文或英文内容，不得出现其他语言或拼写体系
    - 本评论对象为二次元美少女 Galgame，请避免套用其他游戏类型的评价框架
    - 保持客观理性，避免主观臆测

    **详细要求**：

    1. 【内容摘要】：
       - 用3-5句完整的话全面概括长评的核心内容
       - 包含玩家的主要观点、情感体验、具体感受
       - 突出评论中的关键细节和独特见解
       - 体现玩家的情感变化轨迹（兴奋→失望→感动等）
       - 如果评论涉及多个方面，都要在摘要中体现

    2. 【优点总结】：
       - 详细列举玩家提到的所有正面评价
       - 可以按一下维度分类，比如剧情亮点、角色魅力、演出效果、音乐表现、情感共鸣等方面进行优点总结
       - 保留玩家的原始表达和具体例子
       - 挖掘隐含的赞美和积极情绪
       - 用完整段落描述，不要简单罗列

    3. 【缺点总结】：
       - 全面归纳玩家提及的所有不足和问题
       - 包括明确批评和隐含不满
       - 涵盖技术问题、剧情缺陷、角色设定、节奏把控等各方面
       - 保持客观描述，不添加个人判断
       - 如有建设性建议也要包含

    4. 【内容标签】：
       - 从评论中提取0-8个具体标签
       - 可以从剧情深度、角色塑造、情感共鸣、演出效果、音乐氛围、世界观构建、玩家体验等相关内容中总结提取的出需要的标签，要用自己的话而不仅仅是套模板
       - 标签的长度以1-5字为合适，尽量以2字和4字为主
       - 标签要准确反映评论重点，避免泛泛而谈

    **输出标准**：
    - 每个字段都要尽可能详细和丰富
    - 宁可内容多一些，也不要遗漏重要信息
    - 只有在评论真正毫无价值时才留空
    - 保持语言自然流畅，避免机械化表达

    JSON格式如下：
    {{
      "summary": "<详细的内容摘要，3-5句完整描述>",
      "pros": "<详尽的优点总结，包含具体细节和玩家感受>",
      "cons": "<全面的缺点总结，涵盖所有提及的问题>",
      "tags": "<标签1,标签2,标签3,标签4,标签5>",

    }}

    玩家长评如下：
    {truncated_text}
    """

    client = ollama.Client(timeout=300)

    try:
        response = client.chat(
            model="llama3.2",
            messages=[
                {
                    'role': 'system',
                    'content': '你是一个只输出有效JSON的游戏分析助手。',
                },
                {
                    'role': 'user',
                    'content': prompt,
                },
            ],
            format='json',
        )

        analysis_result = json.loads(response['message']['content'])

        # 基本校验
        if all(k in analysis_result for k in ['summary', 'pros', 'cons', 'tags']):
            return analysis_result
        else:
            print("Warning: LLM response JSON missing keys.")
            return None

    except Exception as e:
        print(f"An error occurred during review analysis: {e}")
        return None