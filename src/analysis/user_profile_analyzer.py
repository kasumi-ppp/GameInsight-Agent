import ollama
import json
import pandas as pd
from collections import defaultdict, Counter
from typing import List, Dict, Any, Tuple
import re
import os
import time
from datetime import datetime


class UserProfileAnalyzer:
    """用户画像分析器 - 基于标准流程"""
    
    def __init__(self, model_name: str = "llama3.2", batch_size: int = 20):
        self.model_name = model_name
        self.batch_size = batch_size
        self.client = ollama.Client(timeout=300)
        # 创建save目录
        self.save_dir = "save"
        os.makedirs(self.save_dir, exist_ok=True)
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 整合画像文件路径
        self.combined_profile_file = os.path.join(self.save_dir, f"combined_profiles_{self.timestamp}.json")
        
        # 初始化整合画像数据结构
        self.combined_profiles = {
            'session_info': {
                'timestamp': self.timestamp,
                'model': self.model_name,
                'batch_size': self.batch_size,
                'status': 'in_progress'
            },
            'statistics': {
                'total_games': 0,
                'total_reviews': 0,
                'completed_games': 0,
                'total_processing_time': 0
            },
            'game_profiles': {},
            'overall_profile': None,
            'batch_results': []
        }
    
    def load_existing_session(self, excel_file: str) -> bool:
        """加载现有的分析会话"""
        # 查找现有的整合画像文件
        existing_files = []
        for file in os.listdir(self.save_dir):
            if file.startswith("combined_profiles_") and file.endswith(".json"):
                existing_files.append(file)
        
        if not existing_files:
            return False
        
        # 按时间戳排序，选择最新的
        existing_files.sort(reverse=True)
        latest_file = existing_files[0]
        filepath = os.path.join(self.save_dir, latest_file)
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
            
            # 检查是否是同一个数据文件的分析
            if existing_data.get('session_info', {}).get('excel_file') == excel_file:
                self.combined_profiles = existing_data
                self.combined_profile_file = filepath
                print(f"🔄 发现现有会话: {latest_file}")
                print(f"📊 已完成游戏: {existing_data['statistics']['completed_games']}")
                return True
            else:
                print(f"ℹ️ 发现现有会话但数据文件不同，将创建新会话")
                return False
                
        except Exception as e:
            print(f"⚠️ 加载现有会话失败: {e}")
            return False
    
    def save_combined_profiles(self):
        """保存整合画像"""
        with open(self.combined_profile_file, 'w', encoding='utf-8') as f:
            json.dump(self.combined_profiles, f, ensure_ascii=False, indent=2)
    
    def update_session_progress(self, game_name: str = None, overall_completed: bool = False):
        """更新会话进度"""
        if game_name:
            self.combined_profiles['statistics']['completed_games'] += 1
        
        if overall_completed:
            self.combined_profiles['session_info']['status'] = 'completed'
        
        self.save_combined_profiles()
    
    def save_game_profile(self, game_profile: Dict[str, Any]):
        """保存单个游戏画像到整合文件"""
        if game_profile['status'] == 'success':
            game_name = game_profile['game_name']
            
            # 添加到整合画像
            self.combined_profiles['game_profiles'][game_name] = game_profile
            
            # 更新统计信息
            self.combined_profiles['statistics']['total_reviews'] += game_profile['total_reviews']
            self.combined_profiles['statistics']['total_processing_time'] += game_profile.get('processing_time', 0)
            
            # 保存整合文件
            self.save_combined_profiles()
            
            print(f"  💾 游戏画像已保存到整合文件: {os.path.basename(self.combined_profile_file)}")
            return True
        return False
    
    def save_batch_results(self, game_name: str, batch_results: List[Dict[str, Any]]):
        """保存批次分析结果到整合文件"""
        # 添加到整合画像
        self.combined_profiles['batch_results'].extend(batch_results)
        
        # 保存整合文件
        self.save_combined_profiles()
        
        print(f"  💾 批次分析已保存到整合文件")
        return True
    
    def load_and_group_data(self, excel_file: str) -> Dict[str, List[str]]:
        """加载Excel数据并按游戏分组"""
        try:
            df = pd.read_excel(excel_file)
            
            # 检查必需列
            required_columns = ['游戏名称', '长评内容']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"缺少必需列: {missing_columns}")
            
            # 按游戏名称分组
            games_data = defaultdict(list)
            for _, row in df.iterrows():
                game_name = str(row['游戏名称']).strip()
                review_content = str(row['长评内容']).strip()
                if game_name and review_content:
                    games_data[game_name].append(review_content)
            
            print(f"✅ 成功加载 {len(df)} 条评论，分为 {len(games_data)} 款游戏")
            return dict(games_data)
            
        except Exception as e:
            print(f"❌ 加载数据失败: {e}")
            return {}
    
    def split_into_batches(self, reviews: List[str]) -> List[List[str]]:
        """将评论分批"""
        batches = []
        for i in range(0, len(reviews), self.batch_size):
            batch = reviews[i:i + self.batch_size]
            batches.append(batch)
        return batches
    
    def analyze_batch(self, game_name: str, batch_reviews: List[str], batch_num: int) -> Dict[str, Any]:
        """分析单个批次的用户画像"""
        
        # 构建批次提示词
        reviews_text = "\n".join([f"{i+1}. {review}" for i, review in enumerate(batch_reviews)])
        
        prompt = f"""你是一个资深的用户行为分析师。请基于以下若干条游戏用户评论，提取出该游戏用户在这一批评论中的主要偏好特征。

游戏名称：{game_name}
批次：第{batch_num}批（共{len(batch_reviews)}条评论）

请仔细分析每条评论，提取以下信息：
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


【用户画像】
1. 用户整体评价倾向（正面/负面/中立）及具体表现形式
2. 用户关注的核心要素（如剧情、人物、美术、系统、音乐、氛围等）
3. 用户自己给出的标签或关键词（从评论中直接提取）
4. 情绪或心理特征（如情绪波动、对特定元素的敏感度等）
5. 好评要点（用户明确表达的优点）
6. 差评要点（用户明确表达的不满或建议）

【评论】
{reviews_text}

请确保：
- 区分好评和差评内容
- 提取用户自己使用的标签和关键词
- 分析要具体详细，不要泛泛而谈
- 输出要符合中文语法
- 突出评论中的关键细节和独特见解
- 使用规范中文回答"""


        try:
            print(f"    🤖 正在调用 {self.model_name} 分析批次 {batch_num}...")
            start_time = time.time()
            
            response = self.client.chat(
                model=self.model_name,
                messages=[
                    {
                        'role': 'system',
                        'content': '你是用户行为分析师。请仔细分析用户评论，提取好评、差评、用户标签等信息，输出结构化的用户画像分析，使用规范中文。',
                    },
                    {
                        'role': 'user',
                        'content': prompt,
                    },
                ],
            )
            
            end_time = time.time()
            processing_time = end_time - start_time
            print(f"    ⏱️ 批次 {batch_num} 分析完成，耗时 {processing_time:.1f}秒")
            
            content = response['message']['content'].strip()
            
            # 解析结果
            result = {
                'game_name': game_name,
                'batch_num': batch_num,
                'review_count': len(batch_reviews),
                'analysis': content,
                'processing_time': processing_time,
                'status': 'success'
            }
            
            return result
            
        except Exception as e:
            print(f"❌ 批次分析失败: {e}")
            return {
                'game_name': game_name,
                'batch_num': batch_num,
                'review_count': len(batch_reviews),
                'analysis': f"分析失败: {e}",
                'processing_time': 0,
                'status': 'failed'
            }
    
    def aggregate_game_profiles(self, game_name: str, batch_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """聚合单个游戏的所有批次结果"""
        
        # 过滤成功的批次结果
        successful_batches = [result for result in batch_results if result['status'] == 'success']
        
        if not successful_batches:
            return {
                'game_name': game_name,
                'status': 'failed',
                'error': '所有批次分析都失败了'
            }
        
        # 构建聚合提示词
        batch_analyses = []
        for batch in successful_batches:
            batch_analyses.append(f"批次{batch['batch_num']}（{batch['review_count']}条评论）：\n{batch['analysis']}")
        
        aggregated_text = "\n\n".join(batch_analyses)
        
        prompt = f"""以下是来自同一个游戏的用户画像小结，请你整合这些信息，提取出该游戏的全局用户画像。

游戏名称：{game_name}
总评论数：{sum(batch['review_count'] for batch in successful_batches)}

批次分析结果：
{aggregated_text}

请整合以上信息，输出以下JSON格式：

{{
  "summary": "对游戏的整体评价总结，包含用户的主要观点和情感倾向，根据实际内容实事求是地描述",
  "pros": [
    "用户明确表达的优点1",
    "用户明确表达的优点2",
    "用户明确表达的优点3",
    "根据实际内容添加更多优点..."
  ],
  "cons": [
    "用户明确表达的不满或建议1",
    "用户明确表达的不满或建议2",
    "用户明确表达的不满或建议3",
    "根据实际内容添加更多问题..."
  ],
  "tags": [
    "从评论中提取的用户标签1",
    "从评论中提取的用户标签2",
    "从评论中提取的用户标签3",
    "根据实际内容添加更多标签..."
  ]
}}

请确保：
- summary要实事求是，根据用户的实际评价进行总结，不限制长度
- pros要具体，基于用户的实际表达，有多少优点就列多少
- cons要具体，基于用户的实际表达，有多少问题就列多少
- tags要准确，从评论中直接提取，有多少标签就列多少
- 使用规范中文
- 严格按照JSON格式输出
- 不要为了凑数量而编造内容，实事求是即可"""

        try:
            print(f"  🔄 正在聚合 {game_name} 的画像...")
            start_time = time.time()
            
            response = self.client.chat(
                model=self.model_name,
                messages=[
                    {
                        'role': 'system',
                        'content': '你是用户行为分析师。请整合多个批次的分析结果，输出JSON格式的用户画像，包含summary、pros、cons、tags字段，根据实际内容实事求是地输出，不限制长度，使用规范中文。',
                    },
                    {
                        'role': 'user',
                        'content': prompt,
                    },
                ],
                format='json',
            )
            
            end_time = time.time()
            processing_time = end_time - start_time
            print(f"  ⏱️ {game_name} 聚合完成，耗时 {processing_time:.1f}秒")
            
            content = response['message']['content'].strip()
            
            # 解析JSON结果
            try:
                # 清理可能的markdown标记
                json_match = re.search(r'```(?:json)?\n?(.*?)\n?```', content, re.DOTALL)
                if json_match:
                    content = json_match.group(1).strip()
                
                profile_data = json.loads(content)
                
                # 验证必要字段
                required_fields = ['summary', 'pros', 'cons', 'tags']
                for field in required_fields:
                    if field not in profile_data:
                        profile_data[field] = [] if field in ['pros', 'cons', 'tags'] else ""
                
                return {
                    'game_name': game_name,
                    'total_reviews': sum(batch['review_count'] for batch in successful_batches),
                    'batch_count': len(successful_batches),
                    'summary': profile_data.get('summary', ''),
                    'pros': profile_data.get('pros', []),
                    'cons': profile_data.get('cons', []),
                    'tags': profile_data.get('tags', []),
                    'processing_time': processing_time,
                    'status': 'success'
                }
                
            except json.JSONDecodeError as e:
                print(f"⚠️ JSON解析失败，使用备用格式: {e}")
                return {
                    'game_name': game_name,
                    'total_reviews': sum(batch['review_count'] for batch in successful_batches),
                    'batch_count': len(successful_batches),
                    'summary': f"基于{sum(batch['review_count'] for batch in successful_batches)}条评论的游戏画像",
                    'pros': ["用户评价积极"],
                    'cons': ["需要更多反馈"],
                    'tags': ["用户评价", "游戏体验"],
                    'processing_time': processing_time,
                    'status': 'success'
                }
            
        except Exception as e:
            print(f"❌ 游戏聚合失败: {e}")
            return {
                'game_name': game_name,
                'status': 'failed',
                'error': f'聚合失败: {e}'
            }
    
    def generate_overall_profile(self, game_profiles: List[Dict[str, Any]]) -> Dict[str, Any]:
        """生成整体用户画像"""
        
        # 过滤成功的游戏画像
        successful_profiles = [profile for profile in game_profiles if profile['status'] == 'success']
        
        if not successful_profiles:
            return {
                'status': 'failed',
                'error': '没有成功的游戏画像可聚合'
            }
        
        # 构建整体聚合提示词
        profiles_text = []
        for profile in successful_profiles:
            profile_summary = f"游戏：{profile['game_name']}（{profile['total_reviews']}条评论）\n"
            profile_summary += f"总结：{profile.get('summary', '')}\n"
            profile_summary += f"优点：{', '.join(profile.get('pros', []))}\n"
            profile_summary += f"缺点：{', '.join(profile.get('cons', []))}\n"
            profile_summary += f"标签：{', '.join(profile.get('tags', []))}"
            profiles_text.append(profile_summary)
        
        aggregated_text = "\n\n".join(profiles_text)
        
        prompt = f"""以下是多个不同游戏的用户画像，请你分析所有游戏用户的共性与差异，构建整体用户画像。

游戏数量：{len(successful_profiles)}
总评论数：{sum(profile['total_reviews'] for profile in successful_profiles)}

各游戏用户画像：
{aggregated_text}

请分析所有游戏用户的共性与差异，输出以下JSON格式的整体用户画像：

{{
  "summary": "整体用户画像总结，包含用户偏好、行为特征等，根据实际内容实事求是地描述",
  "pros": [
    "不同游戏用户普遍认可的优点1",
    "不同游戏用户普遍认可的优点2",
    "不同游戏用户普遍认可的优点3",
    "根据实际内容添加更多优点..."
  ],
  "cons": [
    "不同游戏用户普遍反映的问题1",
    "不同游戏用户普遍反映的问题2",
    "不同游戏用户普遍反映的问题3",
    "根据实际内容添加更多问题..."
  ],
  "tags": [
    "高频用户标签1",
    "高频用户标签2",
    "高频用户标签3",
    "根据实际内容添加更多标签..."
  ]
}}

请确保：
- 分析不同游戏类型的用户偏好差异
- 总结用户标签的使用规律
- 分析好评和差评的共性
- summary要实事求是，根据实际内容进行总结，不限制长度
- pros、cons、tags要根据实际内容输出，有多少就列多少
- 不要为了凑数量而编造内容，实事求是即可
- 使用规范中文
- 严格按照JSON格式输出"""

        try:
            print(f"🌐 正在生成整体用户画像...")
            start_time = time.time()
            
            response = self.client.chat(
                model=self.model_name,
                messages=[
                    {
                        'role': 'system',
                        'content': '你是用户行为分析师。请分析多个游戏的用户画像，构建JSON格式的整体用户画像，包含summary、pros、cons、tags字段，根据实际内容实事求是地输出，不限制长度，使用规范中文。',
                    },
                    {
                        'role': 'user',
                        'content': prompt,
                    },
                ],
                format='json',
            )
            
            end_time = time.time()
            processing_time = end_time - start_time
            print(f"⏱️ 整体画像生成完成，耗时 {processing_time:.1f}秒")
            
            content = response['message']['content'].strip()
            
            # 解析JSON结果
            try:
                # 清理可能的markdown标记
                json_match = re.search(r'```(?:json)?\n?(.*?)\n?```', content, re.DOTALL)
                if json_match:
                    content = json_match.group(1).strip()
                
                profile_data = json.loads(content)
                
                # 验证必要字段
                required_fields = ['summary', 'pros', 'cons', 'tags']
                for field in required_fields:
                    if field not in profile_data:
                        profile_data[field] = [] if field in ['pros', 'cons', 'tags'] else ""
                
                return {
                    'total_games': len(successful_profiles),
                    'total_reviews': sum(profile['total_reviews'] for profile in successful_profiles),
                    'summary': profile_data.get('summary', ''),
                    'pros': profile_data.get('pros', []),
                    'cons': profile_data.get('cons', []),
                    'tags': profile_data.get('tags', []),
                    'processing_time': processing_time,
                    'status': 'success'
                }
                
            except json.JSONDecodeError as e:
                print(f"⚠️ JSON解析失败，使用备用格式: {e}")
                return {
                    'total_games': len(successful_profiles),
                    'total_reviews': sum(profile['total_reviews'] for profile in successful_profiles),
                    'summary': f"基于{len(successful_profiles)}款游戏，{sum(profile['total_reviews'] for profile in successful_profiles)}条评论的整体用户画像",
                    'pros': ["用户评价积极"],
                    'cons': ["需要更多反馈"],
                    'tags': ["用户评价", "游戏体验"],
                    'processing_time': processing_time,
                    'status': 'success'
                }
            
        except Exception as e:
            print(f"❌ 整体画像生成失败: {e}")
            return {
                'status': 'failed',
                'error': f'整体画像生成失败: {e}'
            }
    
    def analyze_user_profiles(self, excel_file: str) -> Dict[str, Any]:
        """执行完整的用户画像分析流程"""
        
        print(f"🔄 开始用户画像分析...")
        print(f"📁 数据文件: {excel_file}")
        print(f"🤖 使用模型: {self.model_name}")
        print(f"📦 批次大小: {self.batch_size}")
        print(f"💾 结果将保存到: {self.save_dir}/")
        
        # 检查是否有现有会话
        session_loaded = self.load_existing_session(excel_file)
        if session_loaded:
            print(f"🔄 继续现有会话...")
            # 更新会话信息
            self.combined_profiles['session_info']['excel_file'] = excel_file
            self.save_combined_profiles()
        else:
            print(f"🆕 创建新会话...")
            # 初始化新会话
            self.combined_profiles['session_info']['excel_file'] = excel_file
            self.save_combined_profiles()
        
        # Step 1: 加载和分组数据
        games_data = self.load_and_group_data(excel_file)
        if not games_data:
            return {'status': 'failed', 'error': '数据加载失败'}
        
        # 更新总游戏数
        self.combined_profiles['statistics']['total_games'] = len(games_data)
        self.save_combined_profiles()
        
        # 获取已完成的游戏
        completed_games = set(self.combined_profiles['game_profiles'].keys())
        remaining_games = [game for game in games_data.keys() if game not in completed_games]
        
        if completed_games:
            print(f"📋 已完成的游戏: {', '.join(completed_games)}")
            print(f"📋 待处理的游戏: {', '.join(remaining_games)}")
        
        all_batch_results = self.combined_profiles['batch_results']
        game_profiles = list(self.combined_profiles['game_profiles'].values())
        total_processing_time = self.combined_profiles['statistics']['total_processing_time']
        
        # Step 2: 分批分析每个游戏
        for game_name in remaining_games:
            reviews = games_data[game_name]
            print(f"\n🎮 分析游戏: {game_name} ({len(reviews)}条评论)")
            
            # 分批
            batches = self.split_into_batches(reviews)
            print(f"📦 分为 {len(batches)} 个批次")
            
            batch_results = []
            for i, batch in enumerate(batches, 1):
                print(f"  📋 分析批次 {i}/{len(batches)} ({len(batch)}条评论)")
                result = self.analyze_batch(game_name, batch, i)
                batch_results.append(result)
                all_batch_results.append(result)
                total_processing_time += result.get('processing_time', 0)
            
            # 立即保存批次分析结果
            self.save_batch_results(game_name, batch_results)
            
            # Step 3: 聚合单个游戏的结果
            print(f"  🔄 聚合游戏画像...")
            game_profile = self.aggregate_game_profiles(game_name, batch_results)
            game_profiles.append(game_profile)
            total_processing_time += game_profile.get('processing_time', 0)
            
            # 立即保存游戏画像
            self.save_game_profile(game_profile)
            
            # 更新进度
            self.update_session_progress(game_name)
            
            if game_profile['status'] == 'success':
                print(f"  ✅ {game_name} 画像生成成功")
            else:
                print(f"  ❌ {game_name} 画像生成失败")
        
        # Step 4: 生成整体用户画像
        if not self.combined_profiles['overall_profile']:
            print(f"\n🌐 生成整体用户画像...")
            overall_profile = self.generate_overall_profile(game_profiles)
            total_processing_time += overall_profile.get('processing_time', 0)
            
            # 保存整体画像
            if overall_profile['status'] == 'success':
                self.combined_profiles['overall_profile'] = overall_profile
                self.update_session_progress(overall_completed=True)
                print(f"💾 整体画像已保存到整合文件")
        else:
            print(f"\n🌐 整体画像已存在，跳过生成")
            overall_profile = self.combined_profiles['overall_profile']
        
        # 更新最终统计
        self.combined_profiles['statistics']['total_processing_time'] = total_processing_time
        self.save_combined_profiles()
        
        # 整理结果
        result = {
            'status': 'success',
            'config': {
                'model': self.model_name,
                'batch_size': self.batch_size,
                'excel_file': excel_file
            },
            'statistics': {
                'total_games': len(games_data),
                'total_reviews': sum(len(reviews) for reviews in games_data.values()),
                'total_batches': len(all_batch_results),
                'successful_batches': len([r for r in all_batch_results if r['status'] == 'success']),
                'successful_games': len([p for p in game_profiles if p['status'] == 'success']),
                'total_processing_time': total_processing_time
            },
            'batch_results': all_batch_results,
            'game_profiles': game_profiles,
            'overall_profile': overall_profile,
            'combined_file': self.combined_profile_file
        }
        
        print(f"\n🎉 分析完成！")
        print(f"📊 统计信息:")
        print(f"  - 游戏数量: {result['statistics']['total_games']}")
        print(f"  - 总评论数: {result['statistics']['total_reviews']}")
        print(f"  - 总批次数: {result['statistics']['total_batches']}")
        print(f"  - 成功批次: {result['statistics']['successful_batches']}")
        print(f"  - 成功游戏: {result['statistics']['successful_games']}")
        print(f"  - 总耗时: {total_processing_time:.1f}秒")
        print(f"💾 整合文件: {os.path.basename(self.combined_profile_file)}")
        
        return result 