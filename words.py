import openpyxl
import random
import time
import os

def load_vocabulary(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    vocabulary = [(sheet.cell(row=row, column=1).value, sheet.cell(row=row, column=2).value, sheet.cell(row=row, column=3).value) for row in range(2, sheet.max_row + 1)]
    return vocabulary

def save_to_excel(data, file_path):
    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet["A1"] = "단어"
        sheet["B1"] = "뜻"
        sheet["C1"] = "카운트"

    for idx, (word, meaning, count) in enumerate(data, start=sheet.max_row + 1):
        sheet.cell(row=idx, column=1).value = word
        sheet.cell(row=idx, column=2).value = meaning
        sheet.cell(row=idx, column=3).value = count  # 엑셀 파일에 카운트 값 저장

    wb.save(file_path)

def brain(vocabulary, file_path):
    random.seed(time.time())  # 현재 시간을 시드로 사용
    shuffled_vocabulary = vocabulary[:]  # 복사본 생성
    random.shuffle(shuffled_vocabulary)  # 단어 순서를 무작위로 섞음
    current_index = 0  # 현재 인덱스 초기화
    
    while current_index < len(shuffled_vocabulary):
        word, meaning, count = shuffled_vocabulary[current_index]  # 현재 인덱스의 단어 선택
        
        print("-------------------------")
        print("||단어||:", word)
        print("-------------------------")  
        input("Enter")  # 사용자가 엔터를 입력하면 뜻이 출력됨
        
        print("\n")
        print("-------------------------")
        print("||뜻||:", meaning)
        print("-------------------------")
        
        option = input("Enter 'y' for next word, 'p' for previous word, or type '!exit' to go back to mode selection: ")
        
        if option.lower() == 'y':
            # count + 1
            count += 1
            # 수정된 카운트를 원래의 vocabulary에 반영
            for idx, (w, m, c) in enumerate(vocabulary):
                if w == word:
                    vocabulary[idx] = (w, m, count)
                    
            current_index += 1  # 다음 단어로 인덱스 이동
        elif option.lower() == 'p':
            current_index = max(0, current_index - 1)  # 이전 단어로 인덱스 이동
        elif option.lower() == '!exit':
            print("Exiting brain mode...")
            # 현재 카운트 값을 엑셀 파일에 저장
            save_to_excel(vocabulary, file_path)
            break
        else:
            print("Invalid option! Please try again.")

    return vocabulary



def match(vocabulary):
    random.seed(time.time())
    while True:
        pair = random.choice(vocabulary)
        print("*****뜻을 맞추세요:", pair[0])
        answer = input("****[Enter!]: ")
        if answer.lower() == pair[1].lower():
            print("-------------------------")
            print("****정답입니다!")
            print("-------------------------")
            print("\n")
        elif answer.lower() == "!exit":
            break
        else:
            print("-------------------------")
            print("****오답입니다.")
            print("-------------------------")
        
        #단어와 비교해서 빠진 부분 추가해 나타나게

def main():
    file_path = "C:\\Words_PY\\words_Ex.xlsx"  # 파일 경로 설정
    while True:
        mode = input("***   data, brain, match, !exit  *** ")
        if mode == "data":
            # 데이터 입력 모드
            data = []
            while True:
                print("-------------------------")
                word = input("***단어를 입력하세요 or exitonnow***    ")
                print("-------------------------")
                if word == "exitonnow":
                    break
                print("-------------------------")
                meaning = input("***뜻***       ")
                print("-------------------------")
                if word.strip() == "" or meaning.strip() == "":
                    print("***null***")
                    continue
                data.append((word, meaning, 0))  # 카운트 0으로 초기화
            if data:
                save_to_excel(data, file_path)
                print("\n")
                print("********데이터가 저장되었습니다*******")
                print("\n")
                
        elif mode == "brain":
            # 암기 모드
            vocabulary = load_vocabulary(file_path)
            print("총", len(vocabulary), "개의 단어가 로드되었습니다.")
            input("암기 시작하려면 [Enter]")
            vocabulary = brain(vocabulary, file_path)  # file_path를 함께 전달
            save_to_excel(vocabulary, file_path)  # 엑셀 파일에 업데이트된 카운트 저장
            
        elif mode == "match":
            # 맞추기 모드
            vocabulary = load_vocabulary(file_path)
            print("총", len(vocabulary), "개의 단어가 로드되었습니다.")
            input("맞추기 퀴즈를 시작하려면 [Enter]를 누르세요.")
            match(vocabulary)
        elif mode == "!exit":
            print("프로그램을 종료합니다.")
            break
        else:
            print("올바른 모드를 선택하세요.")

if __name__ == "__main__":
    main()
