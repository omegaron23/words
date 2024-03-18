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
  

    wb.save(file_path)



def brain(vocabulary):
    random.seed(time.time())  # 현재 시간을 시드로 사용
    shuffled_vocabulary = vocabulary[:]  # 복사본 생성
    random.shuffle(shuffled_vocabulary)  # 단어 순서를 무작위로 섞음
    for idx, (word, meaning, count) in enumerate(shuffled_vocabulary):
        print("-------------------------")
        print("||단어||:", word)
        print("-------------------------")  
        input("Enter")  # 사용자가 엔터를 입력하면 뜻이 출력됨
        
        print("\n")
        print("-------------------------")
        print("||뜻||:", meaning)
        print("-------------------------")
        input("E")  # 사용자가 엔터를 입력하면 다음 단어로 넘어감
        print("\n")
        
        
    
    print("That was last one")
    exit()

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

def main():
    while True:
        mode = input("*******data, brain, match, !exit***** ")
        if mode == "data":
            file_path = "C:\\Words_PY\\words_Ex.xlsx"
            data = []
            while True:
                print("\n")
                print("-------------------------")
                word = input("***단어를 입력하세요 or exitonnow***        ")
                print("-------------------------")
                print("\n")
                if word == "exitonnow":
                    break
                print("\n")
                print("-------------------------")
                meaning = input("***뜻***          ")
                print("-------------------------")
                print("\n")
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
            file_path = "C:\\Words_PY\\words_Ex.xlsx"
            vocabulary = load_vocabulary(file_path)
            print("총", len(vocabulary), "개의 단어가 로드되었습니다.")
            input("암기 시작하려면 [Enter]")
            vocabulary = brain(vocabulary)
            save_to_excel(vocabulary, file_path)  # 엑셀 파일에 업데이트된 카운트 저장
        elif mode == "match":
            file_path = "C:\\Words_PY\\words_Ex.xlsx"
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
