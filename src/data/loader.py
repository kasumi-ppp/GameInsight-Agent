import pandas as pd
import os

def load_reviews(file_path):
    """
    Loads reviews from the specified Excel file.

    Args:
        file_path (str): The path to the .xlsx file.

    Returns:
        pandas.DataFrame: A DataFrame containing the review data,
                          or None if the file is not found or an error occurs.
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return None

    try:
        df = pd.read_excel(file_path)
        required_columns = ['游戏名称', '长评内容']
        if all(col in df.columns for col in required_columns):
            reviews_df = df[required_columns].dropna(subset=['长评内容'])
            print(f"Successfully loaded {len(reviews_df)} reviews.")
            return reviews_df
        else:
            print("Error: The Excel file does not contain the required columns: '游戏名称' and '长评内容'.")
            return None
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return None
