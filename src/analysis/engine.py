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
    你是一位专业的游戏评论分析师，擅长从玩家长评中提取内容要点、情绪倾向与核心要素。请仔细阅读以下玩家长评，完成以下任务：

    1. 【内容摘要】：用简洁自然语言总结这条长评的核心观点与情感体验，涵盖主要情节描述、角色评价和玩家共鸣，控制在200字以内；
    2. 【优点总结】：从剧情、角色塑造、演出表现、情感深度等维度，归纳玩家认为的主要优点，不要机械列点，而要用完整语句概括总结；
    3. 【缺点总结】：指出玩家提到的主要问题或槽点，涵盖逻辑问题、节奏问题、角色崩坏等，同样避免列点，采用自然语言描述；
    4. 【内容标签】：基于该长评本身提及的要素，从“剧情”“角色”“氛围”“演出”“音乐”“哲学主题”“世界观构建”“玩家情感共鸣”等维度中选择最相关的3-5个标签，必须具体明确，不允许使用“吐槽”“抽象”“感想”类泛泛词汇；
    5. 【标签属性分类】：为每个标签指定对应的内容主题分类（例如：剧情->叙事，角色->角色，演出->演出等），输出为一个键值对形式。

    请以**严格的JSON格式**输出，结构如下：
    {{
      "summary": "<内容摘要>",
      "pros": "<优点总结>",
      "cons": "<缺点总结>",
      "tags": "<标签1,标签2,标签3>",
      "tag_attributes": {{
        "标签1": "所属维度",
        "标签2": "所属维度",
        ...
      }}
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
        if all(k in analysis_result for k in ['summary', 'pros', 'cons', 'tags', 'tag_attributes']):
            return analysis_result
        else:
            print("Warning: LLM response JSON missing keys.")
            return None

    except Exception as e:
        print(f"An error occurred during review analysis: {e}")
        return None