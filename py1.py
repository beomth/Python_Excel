#파이썬에서 VBA 실행시키기 With py1.xlsm

import xlwings as xw
#엑셀 매크로파일 열기(path는 매크로 파일이 있는 경로)
path = r"C:/Users/ASUS/Desktop/Python/"
wb = xw.Book(path + "py1.xlsm")

#엑셀 VBA의 매크로 함수 'test'를 파이썬 함수로 지정
macro_test = wb.macro('test')

#VBA함수 실행
macro_test()

#함수를 실행한 엑셀파일 따로 저장하기
wb.save("py1_result.xlsm")

#Workbook 객체 닫기
wb.close()

wb = xw.Book(path + "py1.xlsm")
wb.app.visible = True  # 엑셀 창을 사용자에게 보이게
