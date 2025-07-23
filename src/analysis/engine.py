import ollama
import json

def analyze_review_pros_cons(review_text: str) -> dict | None:
    """
    Analyzes a single game review to extract pros and cons using an LLM.

    Args:
        review_text (str): The text of the game review.

    Returns:
        dict: A dictionary containing 'pros' and 'cons' lists,
              or None if analysis fails.
    """
    # Truncate the review to avoid exceeding context limits or slow processing.
    # 2000 characters are usually enough to capture the main points.
    truncated_text = review_text[:2000]

    prompt = f"""
    You are a professional game review analyst. Read the following game review carefully.
    Your task is to extract the key positive points (Pros) and negative points (Cons) mentioned by the player.
    - List up to 3 main Pros.
    - List up to 3 main Cons.
    - If a point is not clearly mentioned, do not include it in the list.
    - Your response MUST be a valid JSON object. Do not add any text before or after the JSON.
    - The JSON object should look like this: {{"pros": ["..."], "cons": ["..."]}}

    Review text:
    ---
    {truncated_text}
    ---
    """

    # Create a client with a longer timeout for the first request
    client = ollama.Client(timeout=300)

    try:
        response = client.chat(
            model="llama3.2",
            messages=[
                {
                    'role': 'system',
                    'content': 'You are a game analysis assistant that responds only with valid JSON.',
                },
                {
                    'role': 'user',
                    'content': prompt,
                },
            ],
            format='json',  # Ask Ollama to ensure the output is JSON
        )

        # The 'format="json"' parameter ensures the content is a JSON object.
        analysis_result = json.loads(response['message']['content'])
        
        # Basic validation
        if 'pros' in analysis_result and 'cons' in analysis_result:
            return analysis_result
        else:
            print("Warning: LLM response JSON is missing 'pros' or 'cons' keys.")
            return None

    except Exception as e:
        print(f"An error occurred during review analysis: {e}")
        return None
