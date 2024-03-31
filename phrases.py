import requests

def check_grammar(text):
    url = "https://api.languagetool.org/v2/check"
    params = {
        "text": text,
        "language": "en-US"  # 언어 설정 (영어)
    }
    response = requests.get(url, params=params)
    data = response.json()
    matches = data.get('matches', [])
    
    if matches:
        print("문법 오류가 발견되었습니다:")
        for match in matches:
            message = match.get('message', '')
            print("- " + message)
    else:
        print("문법 오류가 발견되지 않았습니다.")

# 계속해서 문장을 입력받고 문법 검사를 수행하는 루프
while True:
    # 사용자로부터 문장 입력 받기
    sentence = input("문법을 검사할 문장을 입력하세요 ('!exit'을 입력하면 종료됩니다): ")
    
    # 종료 조건 확인
    if sentence.lower() == '!exit':
        print("프로그램을 종료합니다.")
        break
    
    # 문법 검사 수행
    check_grammar(sentence)
