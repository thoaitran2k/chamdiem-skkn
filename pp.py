import requests

# Thay API_KEY bằng key thật của bạn
API_KEY = "sk-c3add8b055434960b04b13bbc744af57"

headers = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

payload = {
    "model": "deepseek-chat",
    "messages": [{"role": "user", "content": "Trả lời: OK"}],
    "max_tokens": 10
}

try:
    response = requests.post(
        "https://api.deepseek.com/v1/chat/completions",
        headers=headers,
        json=payload,
        timeout=30
    )
    
    print(f"Status code: {response.status_code}")
    
    if response.status_code == 200:
        result = response.json()
        print("✅ API hoạt động tốt!")
        print(f"Response: {result['choices'][0]['message']['content']}")
    else:
        print(f"❌ Lỗi: {response.text}")
        
except Exception as e:
    print(f"❌ Exception: {e}")