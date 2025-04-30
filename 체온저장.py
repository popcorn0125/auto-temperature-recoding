from tkinter import *
import tkinter as tk
import tkinter.font as font
from tkinter import messagebox
from openpyxl import load_workbook
import win32com.client
import win32timezone
import traceback
import sys
from pathlib import Path
import random

global QUIT
QUIT = 0

FILE_PATH = Path.cwd() # *.py가 위치하고 있는 파일 경로

NAME_FILE_PATH = FILE_PATH / "name.xlsx" # 다운받은 파일 경로 
MONITORING_FILE_PATH = FILE_PATH / "코로나19 입소자 일일모니터링.xlsx" # 일일 모니터링 파일 경로  

""" 체온을 저장하는데 필요한 함수 모음"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
####### 체온 저장 #######################################################################
def Temperature(name_arr, date_arr, wb_DF, wb_SF, ws_SF):
    no_name = []
    
    name_arr.clear() # 이름이 저장된 배열 초기화
    date_arr.clear() # 날짜가 저장된 배열 초기화
    
    for i in range(3, 200, 3): # 일일 모니터링에서 이름을 저장하기위한 반복문 
        
        name = ws_SF.Cells(i,1).Value
        
        if (name is None): # 이름이 없으면 중단
            
            break
        
        name_arr.append(name)
        
    for i in range(4, 25, 1): # 일일 모니터링에 날짜들을 저장하기 위한 반복문 

        val = str(ws_SF.Cells(2, i).Value) # str() 넣기
        
        date_arr.append(val)

    now_sheet_num = 0 #현재 시트 번호 (수급자별 파일), 첫번째 시트 부터 시작
     
    sheet_count = len(wb_DF.sheetnames) # 시트의 총 개수 (수급자별 파일)
    
    for i in range(now_sheet_num, sheet_count, 1): # name파일의 시트 갯수 만큼 반복
    
        row_number = 10 # 수급자별에서 날짜가 시작하는 행의 번호
    
        wb_DF.active = i # 현재 시트 활성

        ws_DF = wb_DF.active 

        name = ws_DF.cell(5,3).value # name파일의 5행3열의 값을 변수에 저장
        
        try: # 예외 처리 (이름이 있는지 없는지) - 이름이 있는 경우
            row_num = (name_arr.index(name) + 1) * 3 # 일일모니터링에서의 이름이 해당되는 행 번호
        except Exception : # 이름이 없으면
            no_name.append(name) # 없는 이름을 저장할 no_name 배열에 저장
            wb_SF.Save()
            


        while True: # 체온을 저장하는데 한 명당 최대 3개 까지 체온 입력이 가능하므로 1개일 때, 2개일 때, 3개일 때로 나누어서 처리

            re = ws_DF.cell(row = row_number, column = 1).value # 첫 번째 체온이 저장된 날짜를 저장
        
            re2 = ws_DF.cell(row = row_number+2,column = 1).value # 두 번째 체온을 저장된 날짜를 저장
        
            if re == None: # 체온이 저장된 날짜가 없으므로 반복 멈춤
                break
        
            if re2 == None: # 체온 한 개가 있을 때 (날짜가 같다면 값 none 저장)
            
                value = ws_DF.cell(row = row_number,column = 16).value
            
                value2 = ws_DF.cell(row = row_number+2, column = 16).value
            
                re3 = ws_DF.cell(row = row_number+4, column = 1).value
            
                if re3 == None: # 체온이 2개일 때
                    row_number += 6
                
                else: # 체온이 3개일 때
                    row_number += 4

            else:
            
                value = ws_DF.cell(row = row_number,column = 16).value
            
                value2 = ""
            
                row_number += 2
    
            date = Date_change(re)
        
            try: # 일일 모니터링에 해당 날짜가 있을 경우
                num = date_arr.index(date) + 4
            except Exception : # 해당 날짜가 없을 시 예외 처리
                wb_SF.Save()
                root = Tk()
                root.withdraw()
                messagebox.showwarning("꼭! 꼭! 꼭! 아래의 알림을 읽어주시기 바랍니다.", "일일모니터링의 날짜 범위를 초과하였습니다.\n\n수급자별 날짜를 다시 확인하여 주시기 바랍니다.\n\n확인 후 다시 실행해주시면 됩니다.\n\n")
                root.destroy()
                                       
                
            ws_SF.Cells(row_num,num).Value = value

            ws_SF.Cells(row_num+1,num).Value = value2

    wb_SF.Save()
    #name_print = ' '.join(no_name) # 일일 모니터링에서 없는 이름들을 한꺼번에 출력하기 위한 이어붙이기
    if no_name: # 이름이 한개이상 존재할 경우
        name_print = ' '.join(no_name) # 이름들을 한번에 보이기위해 이어붙이기 
        no_name.clear() # 배열 초기화
        return name_print
    return None

######## 날짜 변환(Temperature 함수에서 실행) ###################
def Date_change(date):
    d = date.split('.')
    d = '-'.join(d)
    d = d + " 00:00:00+00:00"
    return d

####################### 인쇄 여부 출력 ###################################
def PRINT_ALERT(date_arr, no_name_print, ws_DF, wb_SF, ws_SF):

    a = ws_DF.cell(row = 2, column = 1).value

    a = a.split(' ')

    a1 = a[0]
    a1 = a1.split('.')
    a1 = '-'.join(a1)
    a1 = a1 + " 00:00:00+00:00"

    a2 = a[2]
    a2 = a2.split('.')
    a2 = '-'.join(a2)
    a2 = a2 + " 00:00:00+00:00"

    start_date_index = date_arr.index(a1)
    end_date_index = date_arr.index(a2)

    data_count = end_date_index - start_date_index + 1

    num = ws_SF.Cells( 1, 26).Value + data_count

    
    
    if (no_name_print is not None):
        root = Tk()
        root.withdraw()
        messagebox.showwarning("꼭! 꼭! 꼭! 아래의 알림을 읽어주시기 바랍니다.", no_name_print +"\n\n위의 이름이 일일모니터링에 없습니다..\n\n이름을 추가하여 주시기 바랍니다.\n\n이름을 추가 하신 후 다시 실행해 주시면 됩니다.")
        root.destroy()
        return None

    root = Tk()
    root.withdraw()
     
    if num != 21:

        ws_SF.Cells( 1, 26).Value = num

        wb_SF.Save()
        messagebox.showinfo("실행 결과","성공입니다.\n창을 닫으셔도 됩니다.")

    else:

        ws_SF.Cells( 1, 26).Value = 0

        wb_SF.Save()

        messagebox.showinfo("실행 결과", "일일 모니터링 파일을 열어 확인해 보신 후.\n\n이상이 없는 경우 인쇄를 하여 주시기 바랍니다.\n\n감사합니다.^^")

    root.destroy()
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

###################### 체온 삭제 #######################################
def DELETE_Temp(length, wb_SF, ws_SF):

    ra = "D3:X" + str((length * 3 + 2 ))

    ws_SF.Range(ra).ClearContents() # 마지막 행 값은 아직 미정

    wb_SF.Save()

################# 일일 모니터링 날짜 변경 #########################################
def DATE_CHANGE(date, wb_SF, ws_SF):

    arr = date.split(' ')

    arr = arr[0]

    arr = arr.split('-')

    day = int(arr[2]) +1

    if day < 10:

        day = "0" + str(day)
    
    arr[2] = str(day)

    arr = '-'.join(arr)

    ws_SF.Cells(2,4).Value = arr

    wb_SF.Save()
    
    return arr

############ 이름 추가 함수 ################
def Add_Name(name_arr, name, wb, ws): # 이름을 정렬한 다음 추가할 위치에 행을 복사하여 붙여넣기 
    name_arr.append(name)

    name_arr = sorted(name_arr) # 이름 순차적 정렬 

    start_num = name_arr.index(name) # 정렬된 이름들 중에서 해당 이름의 인덱스를 저장

    #end_num = name_arr.index(name_arr[-1]) # 마지막 배열에 저장된 이름 
    end_num = random.randrange(0,len(name_arr))
    while (True):
        if (start_num+1) != end_num:
            break
        end_num = random.randrange(0,len(name_arr))

    
    range1 = 'A'+ str((start_num + 1) * 3)

    range2 = 'X' + str((start_num + 1) * 3 + 2)

    insert_cell_range = range1 + ":" + range2

    range1 = 'D' + str((start_num + 1) * 3) # 붙여넣기 한 데이터를 지우기 위한 cell의 범위

    delete_cell_range = range1 + ":" + range2 # 범위 "D_:X_"

    range1 = 'A'+ str((end_num+1) * 3)

    range2 = 'X' + str((end_num+1) * 3 + 2)

    copy_cell_range = range1 + ":" + range2

    ws.Range(insert_cell_range).Insert() # 3행 삽입 

    ws.Range(copy_cell_range).Copy() # 3행 복사

    ws.Range(insert_cell_range).Select() # 삽입한 행 선택

    ws.Paste() # 선택된 행 붙여넣기

    ws.Range(delete_cell_range).ClearContents() # 붙여넣기한 행의 값들을 삭제(이유: 복사 붙여넣기한 값이므로 새로 추가한 이름의 체온 값이 아니기 때문입니다)
        
    num = (start_num + 1) * 3
        
    ws.Cells(num,1).Value = name

    row_num = len(name_arr)

    height_row_num = row_num * 3
    
    ws.Rows(height_row_num).RowHeight = 25 # 일일 모니터링의 행높이 

    ws.Rows(height_row_num + 1).RowHeight = 25 # 일일 모니터링의 행높이

    ws.Rows(height_row_num + 2).RowHeight = 45 # 일일 모니터링의 행높이


    wb.Save()

    

############### 이름 삭제 함수 안에 반복분 ###############
def Delete_Name( name_arr, del_name_div, wb, ws):
    name_arr = sorted(name_arr) # 이름을 순차적으로 정렬 
    del_name_arr = sorted(del_name_div) # 삭제할 입력받은 이름을 순차적으로 정렬
    count = len(del_name_arr) # 삭제할 이름의 개수

    wrong_name = []
    
    for i in range(count-1, -1, -1): # 삭제하기 위한 반복(삭제할 이름들을 순차적으로 정렬된 이름들 중 마지막 이름부터 삭제 시작. 즉, A,B,C가 있으면 C,B,A순으로 삭제.)

        name = del_name_arr[i]

        try: # 예외 처리 (이름이 있는지 없는지) - 이름이 있는 경우
            num = (name_arr.index(name) + 1) * 3 # 삭제할 이름의 행 번호 저장

            delete_row_range = str(num) + ":" + str(num + 2)

            ws.Rows(delete_row_range).Select()
        
            ws.Rows(delete_row_range).Delete()
        except Exception : # 이름이 없으면
            wrong_name.append(name) # 없는 이름을 저장할 wrong_name 배열에 저장
            
        
        #num = (name_arr.index(name) + 1) * 3

        #delete_row_range = str(num) + ":" + str(num + 2)

        #ws.Rows(delete_row_range).Select()
        
        #ws.Rows(delete_row_range).Delete()

    wb.Save()
    
    if wrong_name is not None:
        return_name = ' '.join(wrong_name)
        return return_name
    return None

######## 입력한 이름 분리 #############
def Div_name(name):

    div_name = name.split(" ")

    return div_name

    
def Temperature_main(): # 체온 저장하기 버튼을 클릭후 실행

    w.destroy() # 창 닫기
    root = Tk()
    root.withdraw()
    messagebox.showinfo("실행결과","실행중이니 조금만 기다려 주세요~")
    root.destroy()
    
    wb_DF = load_workbook(NAME_FILE_PATH)

    ws_DF = wb_DF.active

    excel = win32com.client.Dispatch("Excel.Application")

    wb_SF = excel.Workbooks.Open(MONITORING_FILE_PATH)

    ws_SF = wb_SF.Worksheets("명단")
    
    result = Temperature(name_arr, date_arr, wb_DF, wb_SF, ws_SF) # 체온 저장 함수, 함수 리턴값이 0이면 일일모니터링에 없는 이름이 있고 1이면 이름이 다 포함되어 있음    

    PRINT_ALERT(date_arr, result, ws_DF, wb_SF, ws_SF) # 출력여부 출력 함수

    wb_DF.save(NAME_FILE_PATH)
    #wb_SF.Save()
    wb_DF.close()
    wb_SF.Close(SaveChanges=1)
    excel.Quit()

    
def Delete_and_Change_main(): # 체온 삭제 및 변경 클릭 후 실행

    w.destroy() # 창 닫기
    
    excel = win32com.client.Dispatch("Excel.Application")

    wb_SF = excel.Workbooks.Open(MONITORING_FILE_PATH)

    ws_SF = wb_SF.Worksheets("명단")

    name_arr = []

    date_arr = []

    for i in range(3, 200, 3): # 일일 모니터링에 저장되어있는 날짜들을 저장하기위한 반

        name = ws_SF.Cells(i,1).Value # 일일 모니터링에 있는 이름을 변수에 저장
        
        if (name is None): # 이름이 없으면 멈춤
            
            break
        
        name_arr.append(name) # name변수에 있는 값을 배열에 저장

    for i in range(4, 25, 1):

        val = str(ws_SF.Cells(2, i).Value) # str() 넣기
                
        date_arr.append(val)
    
    end_date = date_arr[20] # 일일 모니터링에서의 마지막 날짜

    length = len(name_arr) # 이름 배열의 총 길이

    DELETE_Temp(length, wb_SF, ws_SF) # 체온 삭제 함수

    text = DATE_CHANGE(end_date, wb_SF, ws_SF) # 날짜 변경 함수

    root1 = Tk()
    root1.withdraw()
    messagebox.showinfo("실행 결과", "날짜가 "+ text + "부터 시작되도록 변경되었습니다.\n\n체온이 삭제되었습니다.\n\n감사합니다^^")
    root1.destroy()

    wb_SF.Save()
    excel.Quit()
    
def Insert_main(): # 이름 추가 또는 삭제 창 실행

    w.destroy() # 창 닫기

    bg_color = '#FFF5E4' 
    
    #sub_win = tk.Toplevel()
    sub_win = Tk()

    sub_win.geometry('650x570')
    
    sub_win.title('이름 추가 또는 삭제')
    sub_win.resizable(False, False)
    sub_win.configure(background = '#FFF5E4')
    
    frame_main = tk.Frame(sub_win, borderwidth=80, width=450, bg = bg_color)
    frame_main.pack(side="top", fill="both")
    
    def Click(value,name): # 확인 버튼을 클릭하면 실행하는 함수
                
        if name=="" and value ==0: # 이름과 추가 또는 삭제를 선택 안했을 경우
            messagebox.showwarning(title = "경고!!", message = "이름 추가 또는 삭제 선택 및 이름을 입력 안 하셔서\n\n할 수 있는 게 아무것도 없습니다. ㅠㅠ", parent = sub_win)
            return
        if name=="": # 이름만 입력 안했을 경우
            messagebox.showwarning(title = "경고!!", message = "이름을 입력해 주시기 바랍니다.", parent = sub_win)
            return
        if value == 0: # 추가 또는 삭제만 선택 안했을 경우
            messagebox.showwarning(title = "경고!!", message = "삭제 또는 추가를 선택해 주시기 바랍니다.", parent = sub_win)
            return
        
        sub_win.destroy() # 창 종료
        
        if value == 1: # 이름 추가
            
            
            excel = win32com.client.Dispatch("Excel.Application")

            wb_SF = excel.Workbooks.Open(MONITORING_FILE_PATH)

            ws_SF = wb_SF.Worksheets("명단")

            NAME = Div_name(name) # 입력한 이름을 쪼개기
            
            count = len(NAME) # 입력한 이름의 개수
            
            for i in range(0, count, 1): # 입력한 이름의 개수만큼 반복

                Add_Name(name_arr, NAME[i], wb_SF, ws_SF) # 이름 추가 함수
                
            wb_SF.Save()
            excel.Quit()
            
            root = Tk()
            root.withdraw()
            messagebox.showinfo("실행결과",name+" 추가되었습니다.\n\n감사합니다^^")
            root.destroy()

        if value == 2: # 이름 삭제
        
            excel = win32com.client.Dispatch("Excel.Application")

            wb_SF = excel.Workbooks.Open(MONITORING_FILE_PATH)

            ws_SF = wb_SF.Worksheets("명단")

            name_arr_new = [] # 이름을 저장할 배열

            for i in range(3, 200, 3): 
                
                name_new = ws_SF.Cells(i,1).Value
                    
                if (name_new is None): # 이름이 없는 경우 반복문 빠져나옴
                        
                    break
                    
                name_arr_new.append(name_new)
                    
            del_name_div = Div_name(name)

            print_wrong_name = Delete_Name(name_arr_new, del_name_div, wb_SF, ws_SF) # 이름 삭제 함수
            
            wb_SF.Save()
            excel.Quit()

            root = Tk()
            root.withdraw()
            #잘못입력된 이름 출력하는 부분
            if print_wrong_name is not None:
                messagebox.showerror("실행결과", "이름 : "+print_wrong_name+"\n\n해당 이름이 존재하지 않거나 잘못 입력되었습니다.\n\n해당 이름을 다시 입력해주시기 바랍니다.\n\n************감사합니다************")
            if print_wrong_name is None:
                messagebox.showinfo("실행결과",name+" 삭제되었습니다.\n\n감사합니다^^")    
            root.destroy()
            
    f = font.Font(size = 22, weight = 'bold')
    
    btn = tk.Button(sub_win, text = "확인", width = 7, relief = "flat", command =lambda: Click(radio_value_type.get(),name_input_entry.get()), bg='light pink', fg = 'white')
    btn['font'] = f
    btn.place(x = 460, y = 20)

    frame_in_del = LabelFrame(frame_main, text=" 추가 or 삭제 선택 ", bg = bg_color)
    frame_in_del['font'] = font.Font(size = 15,)
    frame_in_del.pack(side="top", fill="both", padx=6, pady=10, ipadx=10, ipady=6)
    #frame.place(x = 50, y = 150, width = 200, height = 70)

    
        
    
    radio_value_type = IntVar()
    
    rd = Radiobutton(frame_in_del, text = "이름 추가", variable = radio_value_type, value = 1, width=20, height = 2, bg = bg_color)
    rd.deselect()
    rd['font'] = font.Font(size = 13)
    rd.grid( column = 0, row = 0)

    rd2 = Radiobutton(frame_in_del, text = "이름 삭제", variable = radio_value_type, value = 2, width=20, height = 2, bg = bg_color)
    rd2.deselect()
    rd2['font'] = font.Font(size = 13)
    rd2.grid( column = 1, row = 0)

    frame_name_input = LabelFrame(frame_main, text = " 이름 입력 ", bg = bg_color)
    frame_name_input['font'] = font.Font(size = 15,)
    frame_name_input.pack(side="top", fill="both", padx=6, pady=10, ipadx=10, ipady=6)

    name_label = Label(frame_name_input, width=20, bg = bg_color)
    name_label.pack(side = 'left')

    name_input_label =  Label(name_label, text="이름 : ", bg = bg_color)
    name_input_label['font'] = font.Font(size = 15,)
    name_input_label.grid(column = 0, row=0)
    
    name_input = tk.StringVar()
    name_input_entry = Entry(name_label, textvariable=name_input, width=50)
    name_input_entry.grid(column = 1, row=0, ipady=7)


    text_ex_frame = LabelFrame(frame_main, text = " 이름 적는 예시 ", bg = bg_color)
    text_ex_frame['font'] = font.Font(size = 15, weight='bold')
    text_ex_frame.pack(side="top")

    text_label = Label(text_ex_frame, text = "이름 한 개 적을 경우\n\n 이름:철수\n", bg = bg_color)
    text_label['font'] = font.Font(size = 9)
    text_label.grid(column=0,row=0)

    text_label2 = Label(text_ex_frame, text = "이름 두 개 이상 적을 경우\n\n 이름:짱구 철수 or 이름:짱구 철수 유리 (즉, 이름적고 한칸 띄우기)\n", bg = bg_color)
    text_label2['font'] = font.Font(size = 9)
    text_label2.grid(column=0,row=1)

    text_label3 = Label(text_ex_frame, text = "****중요*****\n\n마지막에는 절대로 한 칸 띄우면 안 됩니다!!!!\n\n", bg = bg_color)
    text_label3['font'] = font.Font(size = 9, weight = 'bold')
    text_label3.grid(column=0,row=2)
    
def Destroy(): # 종료 버튼을 누르면 실행하는 함수
    w.quit()
    global QUIT
    QUIT = 1
    w.destroy() # 종료

def main(w):
    #w = Tk()
        
    w.geometry("980x550")
    w.title("준식's 작품에 오신 걸 환영합니다~ 무엇을 도와드릴까요?")
    w.resizable(False, False) #창 크기 고정(수정 불가능)
    w.configure(background = '#FFE3E1')

    text_bg_color = '#FFF5E4'
    btn_color = '#FF9494'
    text_fg_color = 'white'

    w_frame = tk.Frame(w, width = 400, bg = text_bg_color)
    w_frame.place(x= 3 , y=65 )



    text_frame = LabelFrame(w_frame, text=" 수급자별 파일 다운 받는 방법 및 안내사항 ")
    text_frame['font'] = font.Font(size = 15, weight='bold')
    text_frame.pack(side="top", fill="both", padx=6, pady=10, ipadx=10, ipady=6)    
    text_frame.configure(background = text_bg_color)

    label_font_size = 13

    label = tk.Label(text_frame, text = '1단계:', anchor='w',bg = text_bg_color)
    label['font'] = font.Font(size = label_font_size, weight = 'bold')
    label.grid(column=0,row=0)

    label_1 = tk.Label(text_frame, text = '케어포에 접속 -> 4-5.통합 간호제공 리포트 -> 체온을 저장할 날짜 선택 및 바이탈만 체크-> 조회 버튼 클릭.          ', anchor='w',bg = text_bg_color)
    label_1['font'] = font.Font(size = label_font_size)
    label_1.grid(column=1,row=0)
    space_label = tk.Label(text_frame, text="",bg = text_bg_color)
    space_label.grid(column = 0, row = 1)

    label2 = tk.Label(text_frame, text = '2단계:',anchor='w',bg = text_bg_color)
    label2['font'] = font.Font(size = label_font_size,weight = 'bold')
    label2.grid(column=0,row=2)
    label2_2 = tk.Label(text_frame, text = '수급자를 모두 선택 -> 간호제공 기록지 출력 -> 수급자별 선택 -> 엑셀 저장(왼쪽 상단 초록색이면서 X자 표시)      ', anchor='w',bg = text_bg_color)
    label2_2['font'] = font.Font(size = label_font_size)
    label2_2.grid(column=1,row=2)
    space_label1 = tk.Label(text_frame, text="",bg = text_bg_color)
    space_label1.grid(column = 0, row = 3)

    label3 = tk.Label(text_frame, text = '3단계:',anchor='w',bg = text_bg_color)
    label3['font'] = font.Font(size = label_font_size, weight = 'bold')
    label3.grid(column=0,row=4)

    label3_3 = tk.Label(text_frame, text = '다운받은 파일 열기 -> 편집사용 클릭 -> 파일 -> 정보 -> 호환모드(변환 클릭) -> 알림창이 나오면 예 누르기(두번)',anchor='w',bg = text_bg_color)
    label3_3['font'] = font.Font(size = label_font_size)
    label3_3.grid(column=1,row=4)
    space_label2 = tk.Label(text_frame, text="",bg = text_bg_color)
    space_label2.grid(column = 0, row = 5)

    label4 = tk.Label(text_frame, text = '4단계:',anchor='w',bg = text_bg_color)
    label4['font'] = font.Font(size = label_font_size, weight = 'bold')
    label4.grid(column=0,row=6)

    label4_4 = tk.Label(text_frame, text = '파일 -> 다른 이름으로 저장 -> 찾아보기 -> 바탕화면 -> 파일이름에 name 적기 -> 저장 클릭                             ',anchor='w',bg = text_bg_color)
    label4_4['font'] = font.Font(size = label_font_size)
    label4_4.grid(column=1,row=6)
    space_label3 = tk.Label(text_frame, text="",bg = text_bg_color)
    space_label3.grid(column = 0, row = 7)

    label5 = tk.Label(text_frame, text = '* 4단계까지 완료하셨으면 아래의 버튼을 선택하여 실행하시면 됩니다.',bg = text_bg_color)
    label5['font'] = font.Font(size = label_font_size)
    label5.grid(column=1,row=8)


    label6 = tk.Label(text_frame, text = '* 참고로 현재의 창을 닫지 않고도 일일 모니터링 파일을 실행하실 수 있습니다.', bg = text_bg_color)
    label6['font'] = font.Font(size = label_font_size)
    label6.grid(column=1,row=9)

    label7 = tk.Label(text_frame, text = '**** 감사합니다 ****' , bg = text_bg_color)
    label7['font'] = font.Font(size = label_font_size)
    label7.grid(column=1,row=10)


    for child in text_frame.winfo_children():
        child.grid_configure(padx = 1, pady = 2)

    
    des = Button(w, text="종료", width = 8, height = 1, relief = "flat", command = Destroy, fg = text_fg_color ,bg = btn_color) # Destroy
    des['font'] = font.Font(size = 22, weight='bold')
    des.place(x= 800 , y = 13)

    tem = Button(w, text="체온 저장하기", width = 13, height = 2, relief = "flat", command = Temperature_main, bg = btn_color, fg = text_fg_color)
    tem['font'] = font.Font( size = 22, weight='bold')
    tem.place(x= 40 , y = 430)

    dele= Button(w, text = "날짜 변경 및 체온 삭제",  width = 17, relief = "flat", height = 2, command = Delete_and_Change_main, bg = btn_color, fg = text_fg_color)
    dele['font'] = font.Font(size = 22, weight='bold')
    dele.place(x = 315 , y = 430)

    ins= Button(w, text = "이름 추가 또는 삭제", width = 15, height = 2, relief = "flat", command = Insert_main, bg = btn_color, fg = text_fg_color)
    ins['font'] = font.Font(size = 22, weight='bold')
    ins.place(x = 660 , y = 430)

    w.protocol('WM_DELETE_WINDOW',Destroy) # 창 x버튼을 클릭했을 때 
    
    w.mainloop()


""" 프로그램 실행했을 때 나오는 안내 및 초기값 저장   """

root = Tk()
root.withdraw()
messagebox.showinfo("실행 결과", "실행 중이오니 확인 버튼을 클릭해 주시고\n\n잠시만 기다려 주시기 바랍니다.\n\n감사합니다."+chr(0x2764))
root.destroy()

excel = win32com.client.Dispatch("Excel.Application")

wb_SF = excel.Workbooks.Open(MONITORING_FILE_PATH)

ws_SF = wb_SF.Worksheets("명단")

name_arr = [] # 일일 모니터링의 이름들을 저장

date_arr = [] # 일일 모니터링의 날짜들을 저장

for i in range(3, 200, 3): # 일일 모니터링에 저장되어있는 날짜들을 저장하기위한 반

    name = ws_SF.Cells(i,1).Value # 일일 모니터링에 있는 이름을 변수에 저장
        
    if (name is None): # 이름이 없으면 멈춤
            
        break
        
    name_arr.append(name) # name변수에 있는 값을 배열에 저장

for i in range(4, 25, 1):

    val = str(ws_SF.Cells(2, i).Value) # str() 넣기
            
    date_arr.append(val)
wb_SF.Save()
excel.Quit()

    
while(1):
    w = Tk()
    main(w)
    if QUIT == 1 :
        break
        

    
