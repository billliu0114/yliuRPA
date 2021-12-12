import pyautogui
import time
import xlrd
import pyperclip

# define mouseClick event
# pyautogui doc: https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("Unable to find the matching picture, retry in 0.1s")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 0
        while i < reTry :
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("Repeat")
                i += 1
            time.sleep(0.1)

# Data Check
# cmdType.value  
#   1.0 Left Click
#   2.0 Left Double Click  
#   3.0 Right Click  
#   4.0 Input  
#   5.0 Wait  
#   6.0 Scroll
# ctype     empty: 0
#           string: 1
#           number: 2
#           date: 3
#           boolean: 4
#           error: 5
def dataCheck(sheet1):
    checkCmd = True
    # Check number of rows
    if sheet1.nrows<2:
        print("No Command Found in excel sheet")
        checkCmd = False
    # check data in each row
    i = 1
    while i < sheet1.nrows:
        # col 1, command type check
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('Row ',i+1,", Col 1 has bad data")
            checkCmd = False
        # col 2 Content check
        cmdValue = sheet1.row(i)[1]
        # image reading then click event, need string type data
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('Row ',i+1,", Col 2 has bad data")
                checkCmd = False
        # input event, cannot have blank data
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('Row ',i+1,", Col 2 has bad data")
                checkCmd = False
        # wait event, need number type data
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('Row ',i+1,", Col 2 has bad data")
                checkCmd = False
        # scroll event, content need to be number type
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('Row ',i+1,", Col 2 has bad data")
                checkCmd = False
        i += 1
    return checkCmd

# Robotic process automation
def rpa(img):
    i = 1
    while i < sheet1.nrows:
        # get command type for this row
        cmdType = sheet1.row(i)[0]
        # 1 means single left click
        if cmdType.value == 1.0:
            # get img name
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("Left Click",img)
        # 2 means double left click
        elif cmdType.value == 2.0:
            # get img name
            img = sheet1.row(i)[1].value
            # get retry number of times
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("double click",img)
        # 3 means single right click
        elif cmdType.value == 3.0:
            # get img name
            img = sheet1.row(i)[1].value
            # get retry number of times
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("right click",img) 
        # 4 means input
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("input:",inputValue)                                        
        # 5 means wait
        elif cmdType.value == 5.0:
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("wait ",waitTime," seconds")
        # 6 means scrolling
        elif cmdType.value == 6.0:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("scrolling ",int(scroll)," distance")                      
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # file opening
    wb = xlrd.open_workbook(filename=file)
    # find sheet1
    sheet1 = wb.sheet_by_index(0)
    print('Welcome to YLiu RPA')
    # data check
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        # perform rpa
        rpa(sheet1)  
    else:
        print('Bad command input or program already exits')
