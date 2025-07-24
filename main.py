import os
from dotenv import load_dotenv
from src.data.loader import load_reviews
from src.analysis.engine import analyze_review_full
import json
from tqdm import tqdm
import signal
import sys
from pathlib import Path
from src.data.robust_excel import export_analysis_to_excel_robust, auto_export_on_interrupt_robust

# 全局变量用于信号处理
current_results_file = None


def signal_handler(signum, frame):
    """处理中断信号，自动导出Excel"""
    print(f"\n🛑 接收到中断信号 ({signum})，正在保存当前进度...")
    if current_results_file and Path(current_results_file).exists():
        auto_export_on_interrupt_robust(current_results_file)
    print("✓ 程序已安全退出。")
    sys.exit(0)


def load_existing_results(file_path):
    """Loads existing analysis results from a JSON file."""
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_results(results, file_path):
    """Saves analysis results to a JSON file."""
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=4)


def try_export_excel_robust(results_file_path):
    """尝试使用健壮的Excel导出，不中断主流程"""
    try:
        basic_path, detailed_path = export_analysis_to_excel_robust(results_file_path)
        if basic_path:
            print(f"✓ 基础Excel: {Path(basic_path).name}")
        if detailed_path:
            print(f"✓ 详细Excel: {Path(detailed_path).name}")
    except ImportError as e:
        print(f"⚠️ Excel导出需要安装依赖: {e}")
        print("请运行: pip install pandas openpyxl")
    except Exception as e:
        print(f"⚠️ Excel导出过程中出现错误: {e}")


def main():
    global current_results_file

    # 注册信号处理器
    signal.signal(signal.SIGINT, signal_handler)  # Ctrl+C
    signal.signal(signal.SIGTERM, signal_handler)  # 终止信号

    print("--- Game Insight Agent Initializing ---")
    load_dotenv()

    # --- Phase 1: Data Loading ---
    print("\n[Phase 1: Data Loading]")
    data_file_path = os.path.join("data", "长评数据(1).xlsx")
    reviews_df = load_reviews(data_file_path)
    if reviews_df is None:
        print("\n--- Agent Run Finished (Error in Data Loading) ---")
        return

    # --- Phase 2: Batch Analysis with Checkpointing ---
    print("\n[Phase 2: Batch Analysis]")
    results_file_path = "analysis_results.json"
    current_results_file = results_file_path  # 设置全局变量
    all_results = load_existing_results(results_file_path)

    # Create a set of already processed reviews for quick lookup
    processed_reviews = {res['review_content'] for res in all_results}
    print(f"Found {len(all_results)} existing results. Resuming analysis.")

    # Filter out already processed reviews
    reviews_to_process = reviews_df[~reviews_df['长评内容'].isin(processed_reviews)]

    if reviews_to_process.empty:
        print("All reviews have already been analyzed.")
    else:
        print(f"Starting analysis for {len(reviews_to_process)} new reviews...")

        try:
            # 使用try-except包装主分析循环，正确捕获KeyboardInterrupt
            for index, row in tqdm(reviews_to_process.iterrows(), total=len(reviews_to_process),
                                   desc="Analyzing Reviews"):
                game_name = row['游戏名称']
                review_content = row['长评内容']

                analysis_result = analyze_review_full(review_content)

                if analysis_result:
                    all_results.append({
                        'game_name': game_name,
                        'review_content': review_content,
                        'summary': analysis_result.get('summary', ''),
                        'pros': analysis_result.get('pros', ''),
                        'cons': analysis_result.get('cons', ''),
                        'tags': analysis_result.get('tags', ''),
                        'tag_attributes': analysis_result.get('tag_attributes', {})
                    })

                    # 1. 先保存JSON
                    save_results(all_results, results_file_path)

                    # 2. 再导出Excel
                    print(f"\n✅ Analysis successful for a review of '{game_name}'. Exporting to Excel...")
                    try_export_excel_robust(results_file_path)

                else:
                    print(f"\n❌ Analysis failed for a review of '{game_name}'. Skipping Excel export.")

        except KeyboardInterrupt:
            print(f"\n\n⚠️ 检测到用户中断 (Ctrl+C)，正在保存当前进度...")
            # 保存JSON结果
            save_results(all_results, results_file_path)
            print(f"✓ JSON结果已保存到: {results_file_path}")

            # 导出Excel
            auto_export_on_interrupt_robust(results_file_path)
            print("✓ 程序已安全退出，所有数据已保存。")
            return

        except Exception as e:
            print(f"\n❌ 分析过程中发生意外错误: {e}")
            # 即使出错也要保存当前进度
            save_results(all_results, results_file_path)
            auto_export_on_interrupt_robust(results_file_path)
            raise

    print("\n[Phase 3: Analysis Complete]")
    print(f"Total analyzed reviews in '{results_file_path}': {len(all_results)}")
    print("\n--- Agent Run Finished ---")

    print(f"\n🎉 批量分析完成！")
    print(f"总共处理了 {len(all_results)} 条评论")
    print(f"结果已保存到: {results_file_path}")

    # 最终Excel导出
    print(f"\n📊 正在导出最终Excel文件到save文件夹...")
    try_export_excel_robust(results_file_path)


if __name__ == "__main__":
    main()