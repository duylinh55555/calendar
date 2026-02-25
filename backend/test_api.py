import requests
import json

url = "http://127.0.0.1:5001/api/schedule"

try:
    response = requests.get(url)
    response.raise_for_status()  # Raise an exception for bad status codes (4xx or 5xx)

    # The response from Flask should be UTF-8 encoded
    data = response.json()

    # Pretty print the JSON to a file to be safe
    with open('api_output.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("Successfully fetched data and saved to api_output.json")
    print("First 5 rows:")
    for row in data[:5]:
        print(row)

except requests.exceptions.RequestException as e:
    print(f"An error occurred while making the request: {e}")
except json.JSONDecodeError:
    print("Failed to decode JSON. Response content:")
    print(response.text)
