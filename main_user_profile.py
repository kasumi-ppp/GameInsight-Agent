#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用户画像分析主程序
基于标准流程的用户画像分析系统
"""

import json
import os
import sys
from datetime import datetime

# 添加src目录到路径
sys.path.append('src')

from analysis.user_profile_analyzer import UserProfileAnalyzer


def select_data_file():
    """选择数据文件"""
    data_dir = "data"
    if not os.path.exists(data_dir):
        print(f"❌ 数据目录 '{data_dir}' 不存在")
        return None
    
    # 查找Excel文件
    excel_files = []
    for file in os.listdir(data_dir):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print(f"❌ 在 '{data_dir}' 目录中未找到Excel文件")
        return None
    
    print(f"\n📁 可用的数据文件:")
    for i, file in enumerate(excel_files, 1):
        print(f"  {i}. {file}")
    
    while True:
        try:
            choice = input(f"\n请选择文件 (1-{len(excel_files)}): ").strip()
            choice_num = int(choice)
            if 1 <= choice_num <= len(excel_files):
                selected_file = excel_files[choice_num - 1]
                file_path = os.path.join(data_dir, selected_file)
                print(f"✅ 已选择: {selected_file}")
                return file_path
            else:
                print(f"❌ 请输入 1-{len(excel_files)} 之间的数字")
        except ValueError:
            print("❌ 请输入有效的数字")
        except KeyboardInterrupt:
            print("\n👋 用户取消操作")
            return None


def save_results(results: dict, output_prefix: str = "user_profile_analysis"):
    """保存分析结果到save文件夹"""
    # 创建save目录
    save_dir = "save"
    os.makedirs(save_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 保存完整结果
    full_result_file = os.path.join(save_dir, f"{output_prefix}_{timestamp}.json")
    with open(full_result_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    # 保存游戏画像摘要
    if results.get('game_profiles'):
        game_profiles = {}
        for profile in results['game_profiles']:
            if profile['status'] == 'success':
                game_profiles[profile['game_name']] = {
                    'total_reviews': profile['total_reviews'],
                    'global_profile': profile['global_profile']
                }
        
        summary_file = os.path.join(save_dir, f"game_profiles_{timestamp}.json")
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(game_profiles, f, ensure_ascii=False, indent=2)
    
    # 保存整体画像
    if results.get('overall_profile') and results['overall_profile']['status'] == 'success':
        overall_file = os.path.join(save_dir, f"overall_profile_{timestamp}.json")
        with open(overall_file, 'w', encoding='utf-8') as f:
            json.dump(results['overall_profile'], f, ensure_ascii=False, indent=2)
    
    # 保存批次分析结果
    if results.get('batch_results'):
        batch_file = os.path.join(save_dir, f"batch_analysis_{timestamp}.json")
        with open(batch_file, 'w', encoding='utf-8') as f:
            json.dump(results['batch_results'], f, ensure_ascii=False, indent=2)
    
    return {
        'full_result': full_result_file,
        'game_profiles': summary_file if results.get('game_profiles') else None,
        'overall_profile': overall_file if results.get('overall_profile') else None,
        'batch_analysis': batch_file if results.get('batch_results') else None
    }


def show_menu():
    """显示主菜单"""
    print("\n" + "="*50)
    print("🎮 用户画像分析系统")
    print("="*50)
    print("1. 开始用户画像分析")
    print("2. 查看帮助信息")
    print("3. 退出程序")
    print("="*50)


def show_help():
    """显示帮助信息"""
    print("\n📖 帮助信息:")
    print("="*50)
    print("🎯 功能说明:")
    print("  - 基于Excel文件中的长评数据")
    print("  - 按游戏分组并分批分析")
    print("  - 生成各游戏的用户画像")
    print("  - 构建整体用户画像")
    print("  - 分析好评、差评和用户标签")
    print("  - 断点续传功能")
    print("  - 整合画像保存到单个文件")
    print()
    print("📋 数据格式要求:")
    print("  - Excel文件必须包含 '游戏名称' 列")
    print("  - Excel文件必须包含 '长评内容' 列")
    print("  - 数据文件应放在 'data' 目录中")
    print()
    print("⚙️ 配置参数:")
    print("  - 模型: llama3.2 (默认)")
    print("  - 批次大小: 20条评论/批 (默认)")
    print("  - 超时时间: 300秒")
    print()
    print("📊 输出文件 (保存到save文件夹):")
    print("  - 整合画像文件: combined_profiles_YYYYMMDD_HHMMSS.json")
    print("  - 包含所有游戏画像、批次分析、整体画像")
    print("  - 支持断点续传，中断后可继续处理")
    print("="*50)


def main():
    """主函数"""
    print("🎮 欢迎使用用户画像分析系统！")
    
    while True:
        show_menu()
        
        try:
            choice = input("\n请选择操作 (1-3): ").strip()
            
            if choice == "1":
                # 开始用户画像分析
                print("\n🔄 开始用户画像分析...")
                
                # 选择数据文件
                data_file = select_data_file()
                if not data_file:
                    continue
                
                # 配置分析器
                try:
                    batch_size = input("请输入批次大小 (默认20): ").strip()
                    batch_size = int(batch_size) if batch_size else 20
                except ValueError:
                    print("❌ 批次大小必须是数字，使用默认值20")
                    batch_size = 20
                
                # 创建分析器
                analyzer = UserProfileAnalyzer(
                    model_name="llama3.2",
                    batch_size=batch_size
                )
                
                # 执行分析
                print(f"\n🚀 开始分析...")
                results = analyzer.analyze_user_profiles(data_file)
                
                if results['status'] == 'success':
                    print(f"\n🎉 分析完成！")
                    print(f"💾 整合画像文件: {os.path.basename(results['combined_file'])}")
                    print(f"📁 文件位置: {results['combined_file']}")
                    
                    # 显示整体画像摘要
                    if results.get('overall_profile') and results['overall_profile']['status'] == 'success':
                        print(f"\n🌐 整体用户画像摘要:")
                        print("-" * 40)
                        overall = results['overall_profile']
                        print(f"📊 统计: {overall['total_games']}款游戏, {overall['total_reviews']}条评论")
                        print(f"⏱️ 总耗时: {results['statistics']['total_processing_time']:.1f}秒")
                        print(f"📝 画像内容:")
                        content = overall['overall_profile']
                        print(f"  {content[:500]}..." if len(content) > 500 else f"  {content}")
                    
                    # 显示游戏画像统计
                    if results.get('game_profiles'):
                        successful_games = [p for p in results['game_profiles'] if p['status'] == 'success']
                        print(f"\n🎮 游戏画像统计:")
                        print(f"  - 成功生成: {len(successful_games)} 款游戏")
                        for profile in successful_games:
                            print(f"    • {profile['game_name']} ({profile['total_reviews']}条评论)")
                else:
                    print(f"❌ 分析失败: {results.get('error', '未知错误')}")
                
                input("\n按回车键继续...")
                
            elif choice == "2":
                show_help()
                input("\n按回车键继续...")
                
            elif choice == "3":
                print("👋 感谢使用，再见！")
                break
                
            else:
                print("❌ 请输入有效的选项 (1-3)")
                
        except KeyboardInterrupt:
            print("\n👋 用户取消操作，再见！")
            break
        except Exception as e:
            print(f"❌ 发生错误: {e}")
            input("按回车键继续...")


if __name__ == "__main__":
    main() 