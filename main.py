import os
from dotenv import load_dotenv
from src.data.loader import load_reviews
from src.analysis.engine import analyze_review_pros_cons

import json
from tqdm import tqdm

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

def main():
    """Main function to run the game insight agent."""
    print("--- Game Insight Agent Initializing ---")
    load_dotenv()

    if not os.getenv("E2B_API_KEY"):
        print("Warning: E2B_API_KEY is not set. E2B Sandbox features will be disabled.")

    # --- Phase 1: Data Loading ---
    print("\n[Phase 1: Data Loading]")
    data_file_path = os.path.join("save", "长评数据(1).xlsx")
    reviews_df = load_reviews(data_file_path)
    if reviews_df is None:
        print("\n--- Agent Run Finished (Error in Data Loading) ---")
        return

    # --- Phase 2: Batch Analysis with Checkpointing ---
    print("\n[Phase 2: Batch Analysis]")
    results_file_path = "analysis_results.json"
    all_results = load_existing_results(results_file_path)
    
    # Create a set of already processed reviews for quick lookup
    # Assuming each review content is unique enough to be an identifier
    processed_reviews = {res['review_content'] for res in all_results}
    print(f"Found {len(all_results)} existing results. Resuming analysis.")

    # Filter out already processed reviews
    reviews_to_process = reviews_df[~reviews_df['长评内容'].isin(processed_reviews)]

    if reviews_to_process.empty:
        print("All reviews have already been analyzed.")
    else:
        print(f"Starting analysis for {len(reviews_to_process)} new reviews...")
        # Use tqdm for a progress bar
        for index, row in tqdm(reviews_to_process.iterrows(), total=len(reviews_to_process), desc="Analyzing Reviews"):
            game_name = row['游戏名称']
            review_content = row['长评内容']

            try:
                analysis = analyze_review_pros_cons(review_content)
                if analysis:
                    all_results.append({
                        'game_name': game_name,
                        'review_content': review_content,
                        'analysis': analysis
                    })
                else:
                    # Optionally log failures
                    print(f"\nAnalysis failed for a review of '{game_name}'. Skipping.")

            except Exception as e:
                print(f"\nAn unexpected error occurred during analysis for '{game_name}': {e}")
            
            # Save after each analysis for checkpointing
            save_results(all_results, results_file_path)

    print("\n[Phase 3: Analysis Complete]")
    print(f"Total analyzed reviews in '{results_file_path}': {len(all_results)}")
    print("\n--- Agent Run Finished ---")

if __name__ == "__main__":
    main()