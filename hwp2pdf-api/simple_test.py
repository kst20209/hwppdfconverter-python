# 파일명: simple_test.py
import requests

# API Gateway URL (여러분의 실제 URL로 교체하세요)
api_url = "https://4iitu444ha.execute-api.ap-northeast-2.amazonaws.com/prod/convert"

# 텍스트 파일 읽기
with open("test.txt", "rb") as f:
    file_data = f.read()

# API 호출
response = requests.post(
    api_url,
    data=file_data,
    headers={"Content-Type": "application/octet-stream"}
)

# 결과 출력
print(f"상태 코드: {response.status_code}")
print(f"응답 내용: {response.text}")

if response.status_code == 200:
    print("성공!")
else:
    print("실패...")