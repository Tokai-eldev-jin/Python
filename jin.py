
import global_value as g

import getpass
import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import *
import time
import pyautogui



##os.sep⇒バックスラッシュ

#Hwnd
import subprocess
import win32gui
import win32con
import win32process

g.keytable={}
g.keytable["vk_ctrl"]=0x11
g.keytable["vk_alt"]=0x12
g.keytable["vk_tab"]=0x9
g.keytable["vk_return"]=0xD
g.keytable["vk_back"]=0x8
g.keytable["vk_pause"]=0x13
g.keytable["vk_capital"]=0x14                   ##CapsLockキー
g.keytable["vk_esc"]=0x1B
g.keytable["vk_space"]=0x20
g.keytable["vk_pageup"]=0x21
g.keytable["vk_pagedown"]=0x22
g.keytable["vk_left"]=0x25
g.keytable["vk_right"]=0x27
g.keytable["vk_up"]=0x26
g.keytable["vk_down"]=0x28
g.keytable["vk_delete"]=0x2E
g.keytable["vk_help"]=0x2F
g.keytable["vk_0"]=0x30
g.keytable["vk_1"]=0x31
g.keytable["vk_2"]=0x32
g.keytable["vk_3"]=0x33
g.keytable["vk_4"]=0x34
g.keytable["vk_5"]=0x35
g.keytable["vk_6"]=0x36
g.keytable["vk_7"]=0x37
g.keytable["vk_8"]=0x38
g.keytable["vk_9"]=0x39
g.keytable["vk_numpad0"]=0x60
g.keytable["vk_numpad1"]=0x61
g.keytable["vk_numpad2"]=0x62
g.keytable["vk_numpad3"]=0x63
g.keytable["vk_numpad4"]=0x64
g.keytable["vk_numpad5"]=0x65
g.keytable["vk_numpad6"]=0x66
g.keytable["vk_numpad7"]=0x67
g.keytable["vk_numpad8"]=0x68
g.keytable["vk_numpad9"]=0x69
g.keytable["vk_f1"]=0x70
g.keytable["vk_f2"]=0x71
g.keytable["vk_f3"]=0x72
g.keytable["vk_f4"]=0x73
g.keytable["vk_f5"]=0x74
g.keytable["vk_f6"]=0x75
g.keytable["vk_f7"]=0x76
g.keytable["vk_f8"]=0x77
g.keytable["vk_f9"]=0x78
g.keytable["vk_f10"]=0x79
g.keytable["vk_f11"]=0x7A
g.keytable["vk_f12"]=0x7B
g.keytable["vk_f13"]=0x7C
g.keytable["vk_f14"]=0x7D
g.keytable["vk_f15"]=0x7E
g.keytable["vk_f16"]=0x7F
g.keytable["vk_a"]=0x41
g.keytable["vk_b"]=0x42
g.keytable["vk_c"]=0x43
g.keytable["vk_d"]=0x44
g.keytable["vk_e"]=0x45
g.keytable["vk_f"]=0x46
g.keytable["vk_g"]=0x47
g.keytable["vk_h"]=0x48
g.keytable["vk_i"]=0x49
g.keytable["vk_j"]=0x4A
g.keytable["vk_k"]=0x4B
g.keytable["vk_l"]=0x4C
g.keytable["vk_m"]=0x4D
g.keytable["vk_n"]=0x4E
g.keytable["vk_o"]=0x4F
g.keytable["vk_p"]=0x50
g.keytable["vk_q"]=0x51
g.keytable["vk_r"]=0x52
g.keytable["vk_s"]=0x53
g.keytable["vk_t"]=0x54
g.keytable["vk_u"]=0x55
g.keytable["vk_v"]=0x56
g.keytable["vk_w"]=0x57
g.keytable["vk_x"]=0x58
g.keytable["vk_y"]=0x59
g.keytable["vk_z"]=0x5A
g.keytable["vk_shift"]=0x10
g.keytable["vk_ml"]=0x1
g.keytable["vk_henkan"]=0x1C
g.keytable["vk_muhenkan"]=0x1D
g.keytable["vk_mr"]=0x2
g.keytable["vk_end"]=0x23
g.keytable["vk_home"]=0x24
g.keytable["vk_print"]=0x2A
g.keytable["vk_execute"]=0x2B
g.keytable["vk_snap"]=0x2C
g.keytable["vk_Insert"]=0x2D
g.keytable["vk_mb"]=0x4
g.keytable["vk_win1"]=0x5B
g.keytable["vk_win2"]=0x5C
g.keytable["vk_kake2"]=0x6A
g.keytable["vk_tasu2"]=0x6B
g.keytable["vk_sepa"]=0x6C              ##区切り記号キー
g.keytable["vk_sub"]=0x6D               ##テンキー[-]
g.keytable["vk_dec"]=0x6E               ##テンキー[.]
g.keytable["vk_div"]=0x6F               ##テンキー[/]
g.keytable["vk_num"]=0x90
g.keytable["vk_Scroll"]=0x91
g.keytable["vk_kake1"]=0xBA             ##:*
g.keytable["vk_tasu1"]=0xBB             ##;+
g.keytable["vk_kanma"]=0xBC             ##,<
g.keytable["vk_hiku"]=0xBD              ##-=
g.keytable["vk_ten"]=0xBE               ##.>
g.keytable["vk_que"]=0xBF               ##/?
g.keytable["vk_at"]=0xC0                ##@
g.keytable["vk_kakko1"]=0xDB            ##[{
g.keytable["vk_en"]=0xDC                ##\|
g.keytable["vk_kakko2"]=0xDD            ##]}
g.keytable["vk_kara"]=0xDE              ##^~
g.keytable["vk_sura"]=0xE2              ##＼_
g.keytable["vk_caps"]=0xF0              ##英数 CapsLock
g.keytable["vk_kana"]=0xF2              ##カタカナ ひらがな
g.keytable["vk_han"]=0xF3               ##半角／全角
g.keytable["vk_left_click"]=0x01        ##left_click
g.keytable["vk_right_click"]=0x02        ##left_click

g.vk_ctrl=0x11
g.vk_alt=0x12
g.vk_tab=0x9
g.vk_return=0xD
g.vk_back=0x8
g.vk_pause=0x13
g.vk_capital=0x14
g.vk_esc=0x1B
g.vk_space=0x20
g.vk_pageup=0x21
g.vk_pagedown=0x22
g.vk_left=0x25
g.vk_right=0x27
g.vk_up=0x26
g.vk_down=0x28
g.vk_delete=0x2E
g.vk_help=0x2F
g.vk_0=0x30
g.vk_1=0x31
g.vk_2=0x32
g.vk_3=0x33
g.vk_4=0x34
g.vk_5=0x35
g.vk_6=0x36
g.vk_7=0x37
g.vk_8=0x38
g.vk_9=0x39
g.vk_numpad0=0x60
g.vk_numpad1=0x61
g.vk_numpad2=0x62
g.vk_numpad3=0x63
g.vk_numpad4=0x64
g.vk_numpad5=0x65
g.vk_numpad6=0x66
g.vk_numpad7=0x67
g.vk_numpad8=0x68
g.vk_numpad9=0x69
g.vk_f1=0x70
g.vk_f2=0x71
g.vk_f3=0x72
g.vk_f4=0x73
g.vk_f5=0x74
g.vk_f6=0x75
g.vk_f7=0x76
g.vk_f8=0x77
g.vk_f9=0x78
g.vk_f10=0x79
g.vk_f11=0x7A
g.vk_f12=0x7B
g.vk_f13=0x7C
g.vk_f14=0x7D
g.vk_f15=0x7E
g.vk_f16=0x7F
g.vk_a=0x41
g.vk_b=0x42
g.vk_c=0x43
g.vk_d=0x44
g.vk_e=0x45
g.vk_f=0x46
g.vk_g=0x47
g.vk_h=0x48
g.vk_i=0x49
g.vk_j=0x4A
g.vk_k=0x4B
g.vk_l=0x4C
g.vk_m=0x4D
g.vk_n=0x4E
g.vk_o=0x4F
g.vk_p=0x50
g.vk_q=0x51
g.vk_r=0x52
g.vk_s=0x53
g.vk_t=0x54
g.vk_u=0x55
g.vk_v=0x56
g.vk_w=0x57
g.vk_x=0x58
g.vk_y=0x59
g.vk_z=0x5A
g.vk_shift=0x10
g.vk_ml=0x1
g.vk_henkan=0x1C
g.vk_muhenkan=0x1D
g.vk_mr=0x2
g.vk_end=0x23
g.vk_home=0x24
g.vk_print=0x2A
g.vk_execute=0x2B
g.vk_snap=0x2C
g.vk_Insert=0x2D
g.vk_mb=0x4
g.vk_win1=0x5B
g.vk_win2=0x5C
g.vk_kake2=0x6A
g.vk_tasu2=0x6B
g.vk_sepa=0x6C
g.vk_sub=0x6D
g.vk_dec=0x6E
g.vk_div=0x6F
g.vk_num=0x90
g.vk_Scroll=0x91
g.vk_kake1=0xBA
g.vk_tasu1=0xBB
g.vk_kanma=0xBC
g.vk_hiku=0xBD
g.vk_ten=0xBE
g.vk_que=0xBF
g.vk_at=0xC0
g.vk_kakko1=0xDB
g.vk_en=0xDC
g.vk_kakko2=0xDD
g.vk_kara=0xDE
g.vk_sura=0xE2
g.vk_caps=0xF0
g.vk_kana=0xF2
g.vk_han=0xF3
g.vk_left_click=0x01
g.vk_right_click=0x02




##CHKIMG
import cv2
import numpy as np
from copy import deepcopy
import matplotlib.pyplot as plt
import pyautogui


def show(img):      ##表示関数
    plt.imshow(img, vmin=0, vmax=255)
    plt.show()

def get_screen(region):         ##プレイエリアキャプチャ
    img = pyautogui.screenshot(region=region) #PIL系でキャプチャされる
    return np.array(img, 'uint8') #cv2系で返す

def template_match(img, temp, visible=False):
    ##類似度
    res = cv2.matchTemplate(img, temp, cv2.TM_CCOEFF_NORMED)

    ##当たり座標 [[x0, y0], [x1, y1], ...]
    threshold = 0.95

    pos = np.array(np.where(res>=threshold)).T[:, ::-1]
    ##print(pos)

    ##テンプレートの高さ、幅を取得
    h, w = temp.shape[:2]

    ##枠をつけて表示
    if visible:
        img_copy = deepcopy(img)
        for x, y in pos:
            cv2.rectangle(img_copy, (x, y), (x+w, y+h), (255, 0, 0), 2)

        show(img_copy)

    ##当たりの中心座標を返す
    pos[:, 0] += w // 2
    pos[:, 1] += h // 2
    return pos
##CHKIMG


##OCR
##　　tesseract-ocrインストール    https://github.com/UB-Mannheim/tesseract/wiki
##インストール先　"C:\Users\user_name\AppData\Local\Programs\Tesseract-OCR

def render_doc_text(file_path,val,kind=0,kakudai=1):
    from PIL import Image
    import pyocr
    import pyocr.builders
    import cv2
    import numpy as np

    ## ツール取得
    user_name=getuser()
    ##pyocr.tesseract.TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR/tesseract.exe"
    pyocr.tesseract.TESSERACT_CMD = path_join(r"C:\Users",user_name,r"AppData\Local\Tesseract-OCR/tesseract.exe")
    tools = pyocr.get_available_tools()
    tool = tools[0]

    ## 画像取得
    img = cv2.imread(file_path)

    ## リサイズ
    img_size=img.shape              ##幅、高さ、色
    print(img_size)
    size = (int(img_size[1]*kakudai),int(img_size[0]*kakudai))          #幅　高さ
    img = cv2.resize(img,size,interpolation = cv2.INTER_LINEAR)   ## リサイズ後のサイズを指定
    ##img = cv2.resize(img,size)


    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    ## 必要に応じて画像処理 線を消す
##    ret, img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)
    ret, img = cv2.threshold(img, val, 255, cv2.THRESH_BINARY)
    img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 5)
##    img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 51, 20)


    img = cv2.bitwise_not(img)
    label = cv2.connectedComponentsWithStats(img)
    data = np.delete(label[2], 0, 0)
    new_image = np.zeros((img.shape[0], img.shape[1]))+255
    for i in range(label[0]-1):
        if 0 < data[i][4] < 1000:
            new_image = np.where(label[1] == i+1, 0, new_image)

    ## ret, img = cv2.threshold(img, 80, 255, cv2.THRESH_BINARY)
    cv2.imwrite('sample_edited.png', new_image)
    img = Image.fromarray(new_image)

    ## OCR
    if kind==0:
        builder = pyocr.builders.TextBuilder(tesseract_layout=6)
        result = tool.image_to_string(img, lang="jpn", builder=builder)

        ## 結果から空白文字削除
        data_list = [text for text in result.split('\n') if text.strip()]

        return data_list
    else:
        builder = pyocr.builders.WordBoxBuilder(tesseract_layout=6)
        result = tool.image_to_string(img, lang="jpn", builder=builder)
        draw_rectangle = cv2.imread('sample_edited.png')
        tt=0
        OCR_table={}
        for box in result:
            cv2.rectangle(draw_rectangle, box.position[0], box.position[1], (255, 0, 0), 1)
            g.token_str=str(box)
            t=0
            while g.token_str != "":
                getdata=token(" ",g.token_str)
                if t == 0 :
                    OCR_table[tt]=getdata
                else:
                    getdata1=OCR_table[tt]
                    OCR_table[tt]=getdata1+" "+getdata
                t+=1
            tt+=1

        cv2.imwrite('sample_edited.png', draw_rectangle)
        return OCR_table
##OCR





def main():
    pass







def Hairetu(num1,num2=1):
    '''
    num1            :1次元要素数
    num2            :2次元要素数
    return          :配列
    '''
    table=[[0 for i in range(num2)] for ii in range(num1)]
    return table


def Hairetu_sort(hairetu,kind=""):
    '''
    hairetu         :配列
    kind            :""         昇順
                    :"reverse"  降順
    '''
    if kind=="":
        hairetu2=sorted(hairetu)

    elif kind=="reverse":
        hairetu2=sorted(hairetu,reverse=True)

    return hairetu2


def getid(g_title,g_class=None):
    '''
    getid(タイトル、クラス）
    g_title             :タイトル名
    g_class             :クラス名（省略可）
    '''

    from win32gui import GetWindowText,GetClassName,GetForegroundWindow,GetCursorPos
    ##import win32gui

    def winEnumHandler(hwnd,ctx):
        if win32gui.IsWindowVisible(hwnd):
            ##print (hex(hwnd), win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd))
            g.all_win.append(str(hwnd)+","+win32gui.GetWindowText(hwnd)+","+win32gui.GetClassName(hwnd))

    if g_title=="get_active_win":
        hwnd=GetForegroundWindow()
        g.g_title=GetWindowText(hwnd)
        g.g_class=GetClassName(hwnd)
        print(g.g_title+","+g.g_class)
        return hwnd
    elif g_title=="get_frompoint_win":
        myx,myy=win32gui.GetCursorPos()
        hwnd = win32gui.WindowFromPoint((myx, myy))
        GA_ROOT=2
        hwnd = win32gui.GetAncestor(hwnd, GA_ROOT)
        g.g_title=GetWindowText(hwnd)
        g.g_class=GetClassName(hwnd)
        print(g.g_title+","+g.g_class)
        return hwnd
    elif g_title=="all_win":
        g.all_win=list()
        win32gui.EnumWindows(winEnumHandler,None)
        return g.all_win
    else:
        hwnd = win32gui.FindWindow(g_class,g_title)
        if hwnd != 0:
            g.g_title=GetWindowText(hwnd)
            g.g_class=GetClassName(hwnd)
            print(g.g_title+","+g.g_class)
        return hwnd


def ACW(hwnd,x,y,wx=0,wy=0):
    '''
    hWnd            :ウィンドウハンドル
    x               :頂点座標X
    y               :頂点座標Y
    wx              :幅
    wy              :高さ
    '''

    import win32gui

    if wx==0 or wy==0:
        get_gamen_size()
        wx=g.g_screen_w
        wy=g.g_screen_h

    win32gui.MoveWindow(hwnd, x, y, wx, wy,1)


def ctrlwin(hwnd,meirei):
    '''
    hWnd                :ウィンドウハンドル
    meirei              :"fs_topmost"
                        :"fs_notopmost"
                        :"fs_active"
                        :"fs_show"
                        :"fs_min"
                        :"fs_close"
    '''

    import win32gui
    import win32con
    import win32api
    import pyautogui

    SW_SHOWNORMAL=1
    SW_MAXIMIZE = 3
    WM_CLOSE = 0x10

    if meirei == "fs_topmost":
        #win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetWindowPos(hwnd,-1,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        pyautogui.moveTo(left+60, top + 10)
        pyautogui.click()
    elif meirei == "fs_notopmost":
        win32gui.SetWindowPos(hwnd,-2,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        pyautogui.moveTo(left+60, top + 10)
        pyautogui.click()
    elif meirei == "fs_active":
        win32gui.SetForegroundWindow(hwnd)
    elif meirei == "fs_show":
        win32gui.ShowWindow(hwnd, SW_SHOWNORMAL)
    elif meirei == "fs_max":
        win32gui.ShowWindow(hwnd, SW_MAXIMIZE)
    elif meirei == "fs_min":
        win32gui.CloseWindow(hwnd)
    elif meirei == "fs_close":
        win32gui.SendMessage(hwnd, WM_CLOSE,0,0)


def get_sub_hWnd(p_handle):     ##親ウィンドウから子ウィンドウを取得する
    '''
    get_sub_hWnd(hwnd)
    p_handle                :ハンドル
    '''
    import win32gui
    import win32con
    import array
    import ctypes
    import os.path

    ## 親ウィンドウハンドル(識別番号)とそのクラス名、ハンドル名を取得
    ##p_handle = win32gui.FindWindow(None,"XXX") # アプリケーションのタイトル文字列など入力
    p_class_name = win32gui.GetClassName(p_handle)
    p_handle_name = win32gui.GetWindowText(p_handle)
    handles_dict = {str(p_handle): [p_class_name,p_handle_name]}
    ##print('p',hex(p_handle),p_class_name,p_handle_name)

    ## 親ウィンドウハンドルの子ウィンドウハンドルを配列に格納
    c_handles = array.array("i")  # 子ウィンドウハンドルは複数あることを想定し配列を作成
    win32gui.EnumChildWindows(p_handle,lambda handle,list: list.append(handle),c_handles)

    ## 子ウィンドウハンドルとそのクラス名、ハンドル名を列挙
    g.hWnd_list={}
    g.hWnd_list2={}
    g.hWnd_list3={}

    str_caption=""
    user32 = ctypes.WinDLL("user32")
    get_str=ctypes.create_unicode_buffer(256)

    for hWnd in c_handles:
        c_class_name = win32gui.GetClassName(hWnd)
        c_handle_name = win32gui.GetWindowText(hWnd)
        c_rect = win32gui.GetWindowRect(hWnd)
        rect2=list()
        rect2.append(c_rect[0])
        rect2.append(c_rect[1])
        rect2.append(c_rect[2]-c_rect[0])
        rect2.append(c_rect[3]-c_rect[1])

        length = user32.GetWindowTextLengthW(hWnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hWnd, buff, length + 1)
        strLen = win32gui.SendMessage(hWnd, 0xD, 256, get_str) ##'''''読み取り
        ##print(buff.value,get_str.value)

        print('c',hWnd,c_class_name,"["+c_handle_name+"]","["+get_str.value+"]",rect2)
        handles_dict[hWnd] = [c_class_name,c_handle_name]
        g.hWnd_list2[str(hWnd)+"_"+str(c_class_name)]=str(rect2[0])+"_"+str(rect2[1])+"_"+str(rect2[2])+"_"+str(rect2[3])
        g.hWnd_list3[hWnd]=c_handle_name+"§"+c_class_name

        if hash_data_out(g.hWnd_list,"hash_exists",c_class_name)==False:
            g.hWnd_list[c_class_name]=str(hWnd)
        else:
            getdata=g.hWnd_list[c_class_name]
            g.hWnd_list[c_class_name]=getdata+"_"+str(hWnd)

        for k in g.hWnd_list.keys():
            print(k)
            gettable=g.hWnd_list[k].split("_")
            for p in range(0,len(gettable)):
                print(p+1,gettable[p])


def getstr(hwnd,kind="EDIT",num=1):
    '''
    hwnd            :0                クリップボード
                    :0以外            親ハンドル
    kind            :クラス名
    num             :同じクラス名の何個目か
    '''

    if hwnd != 0:
        import win32gui
        import ctypes

        get_str=ctypes.create_unicode_buffer(256)

        get_sub_hWnd(hwnd)

        get_hwnd=g.hWnd_list[kind].split("_")
        get_hwnd_num=int(get_hwnd[num-1])

        win32gui.SendMessage(get_hwnd_num,0xD,256,get_str)
        return get_str.value

    #クリップボード
    else:
        import pyperclip
        return pyperclip.paste()


def sendstr(hwnd,Sendmoji,kind="EDIT",num=1):
    '''
    hwnd            :0            クリップボード
                    :0以外        親ハンドル
    Sendmoji        :送る文字
    kind            :クラス名
    num             :同じクラス名の何個目か
    '''


    if hwnd != 0:
        import win32gui

        get_sub_hWnd(hwnd)

        get_hwnd=g.hWnd_list[kind].split("_")
        get_hwnd_num=int(get_hwnd[num-1])

        win32gui.SendMessage(get_hwnd_num,0xC, 0, Sendmoji)

    #クリップボード
    else:
        import pyperclip
        pyperclip.copy(Sendmoji)


def clkitem(hwnd,k_class,do_kind,num=1,do_name=""):
    '''
    hwnd            :親ウィンドウハンドル
    k_class         :子ウィンドウのクラス名
    do_kind         :"clk_btn"
                    :"clk_combo_select"
                    :"clk_list_select"
                    :"clk_check"
                    :"combo_count"               項目数を得る
                    :"list_count"
                    :"combo_selectnum"           項目番号を選択する
                    :"list_selectnum"
                    :"combo_dropdown"            ドロップダウンする
                    :"combo_select_value"        選択されているアイテム名
                    :"list_select_value"
                    :"scroll_pos"                ScrollBarのポジションを取得
                    :"scroll_set"                セットされるが、更新されないので、後でバーをクリックする必要あり
                    :"zahyou"                    座標を返す
    num             :同じクラス名の何個目か
    do_name         :select時の選択するアイテム名
    '''

    import win32gui
    import win32con
    import ctypes

    user32 = ctypes.WinDLL("user32")
    get_sub_hWnd(hwnd)
    get_hwnd=g.hWnd_list[k_class].split("_")
    get_hwnd_num=int(get_hwnd[num-1])

    ctrlwin(hwnd,"fs_active")

    if do_kind == "clk_btn":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0xF5, 0, 0)#クリック
        ret=""
    elif do_kind == "combo_count":
        ret=win32gui.SendMessage(get_hwnd_num,win32con.CB_GETCOUNT,0,0)
    elif do_kind == "clk_combo_select":
        win32gui.SendMessage(get_hwnd_num,win32con.CB_SELECTSTRING,0,do_name)  # ドロップダウンを表示する
        ret=""
##        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
##        user32.SendMessageA(get_hwnd_num, 0x14D, 0, do_name) #''''選択
##        user32.PostMessageA(hwnd,0x111,0x10000,get_hwnd_num) #''''選択
    elif do_kind == "combo_selectnum":
        do_name=int(do_name)
        ret=win32gui.SendMessage(get_hwnd_num,win32con.CB_SETCURSEL,do_name,0)
    elif do_kind == "combo_dropdown":
        ret=win32gui.SendMessage(get_hwnd_num,win32con.CB_SHOWDROPDOWN,True,0)
    elif do_kind == "combo_select_value":
        num2=win32gui.SendMessage(get_hwnd_num,win32con.CB_GETCURSEL,0,0)
        size=win32gui.SendMessage(get_hwnd_num,win32con.CB_GETLBTEXTLEN,num2,"")
        buff = ctypes.create_unicode_buffer(size + 1)
        win32gui.SendMessage(get_hwnd_num,win32con.CB_GETLBTEXT,num2,buff)
        ret=buff.value
    elif do_kind == "clk_list_select":
        win32gui.SendMessage(get_hwnd_num,win32con.LB_SELECTSTRING,0,do_name)  # ドロップダウンを表示する
        ret=""
##        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
##        user32.SendMessageA(get_hwnd_num, 0x186, 0, do_name) #''''クリック
    elif do_kind == "list_select_value":
        num2=win32gui.SendMessage(get_hwnd_num,win32con.LB_GETCURSEL,0,0)
        size=win32gui.SendMessage(get_hwnd_num,win32con.LB_GETTEXTLEN,num2,"")
        buff = ctypes.create_unicode_buffer(size + 1)
        win32gui.SendMessage(get_hwnd_num,win32con.LB_GETTEXT,num2,buff)
        ret=buff.value
    elif do_kind == "list_selectnum":
        do_name=int(do_name)
        ret=win32gui.SendMessage(get_hwnd_num,win32con.LB_SETCURSEL,do_name,0)
    elif do_kind == "list_count":
        ret=win32gui.SendMessage(get_hwnd_num,win32con.LB_GETCOUNT,0,0)
    elif do_kind == "clk_check":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0xF0, 0, 0) #''''クリック
    elif do_kind == "scroll_pos":
        ret=win32gui.SendMessage(get_hwnd_num,win32con.SBM_GETPOS,0,0)
    elif do_kind == "scroll_set":
        ret = win32gui.SendMessage(get_hwnd_num,win32con.SBM_SETPOS,do_name,True)
        hwnd_pos=list()
        g.token_str=g.hWnd_list2[str(get_hwnd_num)+"_"+k_class]
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        ret=hwnd_pos
    elif do_kind == "zahyou":
        hwnd_pos=list()
        g.token_str=g.hWnd_list2[str(get_hwnd_num)+"_"+k_class]
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        ret=hwnd_pos

    return ret


def clkitem2(hwnd,do_name,do_kind):
    '''
    hwnd            :親ウィンドウハンドル
    do_name         :アイテム名
    do_kind         :"clk_btn"
                    :"zahyou"                    座標を返す
    '''

    import win32gui
    import win32con
    import ctypes

    user32 = ctypes.WinDLL("user32")
    get_sub_hWnd(hwnd)

    get_table=list()
    for m in range(len(g.hWnd_list3)):
        get_hwnd_num = hash_data_out(g.hWnd_list3,"hash_key",m)
        g.token_str = hash_data_out(g.hWnd_list3,"hash_val",m)
        get_str_value = token("§",g.token_str)
        get_class = token("§",g.token_str)

        if get_str_value == do_name:
            get_table.append(get_str_value+"§"+get_class+"§"+str(get_hwnd_num))

    if len(get_table) == 1:
        g.token_str = get_table[0]
    elif len(get_table) == 0:
        msg("ありません")
        End()
    else:
        getnum=popupmenu2(get_table)
        if getnum < 0 :
            End()
        else:
            g.token_str = get_table[getnum]

    for m in range(3):
        getdata=token("§",g.token_str)
        if m == 1:
            k_class =getdata
        elif m == 2:
            get_hwnd_num=int(getdata)



    ctrlwin(hwnd,"fs_active")

    if do_kind == "clk_btn":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0xF5, 0, 0)#クリック
        ret=""
    elif do_kind == "zahyou":
        hwnd_pos=list()
        g.token_str=g.hWnd_list2[str(get_hwnd_num)+"_"+k_class]
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        hwnd_pos.append(int(token("_",g.token_str)))
        ret=hwnd_pos

    return ret



def get_slider(hwnd,k_class,num=1):
    '''
    hwnd            :ウィンドウハンドル
    k_class         :スクロールバーのクラス
    num             :n個目

    戻り値           :scrollpos[0]       :現在位置
                    :scrollpos[1]       :スクロールバーの最小値
                    :scrollpos[2]       :スクロールバーの最大値
    '''
    import win32gui
    import win32con
    import ctypes

    from ctypes import windll, Structure, c_long, byref

    user32 = ctypes.WinDLL("user32")
    get_sub_hWnd(hwnd)
    get_hwnd=g.hWnd_list[k_class].split("_")
    get_hwnd_num=int(get_hwnd[num-1])

    ctrlwin(hwnd,"fs_active")

    class SLIDER(Structure):
        _fields_ = [
                        ("nMax", c_long),
                        ("nMin", c_long),
                    ]
    SL = SLIDER()

    pos = user32.GetScrollPos(get_hwnd_num,win32con.SB_CTL)
    ret = user32.GetScrollRange(get_hwnd_num,win32con.SB_CTL,byref(SL),byref(SL))

    scrollpos = list(range(3))
    scrollpos[0]=pos
    scrollpos[1]=SL.nMax
    scrollpos[2]=SL.nMin
    print(pos,SL.nMax,SL.nMin)
    return scrollpos


def clkmenu(hwnd,menu_pos=""):
    '''
    メニューをクリックする
    clikmenu(hwnd,"4_3")

    hwnd        :親ウィンドウハンドル
    menu_pos    :クリックするメニューの場所
                :0_0 ファイル/新規
                :0_1 ファイル/新しいウィンドウ
                :1_0 編集/元に戻す
    '''

    menu_hwnd = win32gui.GetMenu(hwnd)  # メニューバーのハンドルを取得
    menu_count = win32gui.GetMenuItemCount(menu_hwnd)  # メニューバー内のサブメニューの個数を取得
    get_list={}

    def get_sub_menu_id(sub_hwnd,sub_menu_count,Astr):
        Bstr=Astr
        for ii in range(0,sub_menu_count):
            get_id=win32gui.GetMenuItemID(sub_hwnd, ii)

            Astr = Bstr + "_" + str(ii)
            if get_id == -1 :
                menu_hwnd2=win32gui.GetSubMenu(sub_hwnd, ii)
                menu_count2=win32gui.GetMenuItemCount(menu_hwnd2)

                get_sub_menu_id(menu_hwnd2,menu_count2,Astr)
            else:
                print(Astr,get_id)
                get_list[Astr]=get_id

    for i in range(0,menu_count):
        sub_hwnd = win32gui.GetSubMenu(menu_hwnd, i)
        sub_menu_count = win32gui.GetMenuItemCount(sub_hwnd)
        Astr=str(i)
        get_sub_menu_id(sub_hwnd,sub_menu_count,Astr)

    if menu_pos != "":
        win32gui.SendMessage(hwnd,win32con.WM_COMMAND,get_list[menu_pos],menu_hwnd)


def status(hwnd):
    '''
    rect=status(hwnd)
    hwnd                :windowハンドル
    戻り値               :rect[0]:画面左上x座標
                        :rect[1]:画面左上y座標
                        :rect[2]:ウィンドウ幅
                        :rect[3]:ウィンドウ高さ
    '''

    import win32gui

    rect = win32gui.GetWindowRect(hwnd)
    rect2=list()
    rect2.append(rect[0])
    rect2.append(rect[1])
    rect2.append(rect[2]-rect[0])
    rect2.append(rect[3]-rect[1])

    return rect2


def get_chr(strA):      ##chrを得る
    num=ord(strA)
    ##print(num,chr(num))
    return num


def getkeystate():
    '''
    押されたキーを返す
    '''
    import ctypes

    def isPressed(key):
        return bool(ctypes.windll.user32.GetAsyncKeyState(key)&0x8000)

    key:str=""
    for k in g.keytable.keys():
        inkey=g.keytable[k]
        getdata=isPressed(inkey)
        if getdata == True:
            if key=="":
                key=k
            else:
                key=key+"@"+k


    return key


def sckey(id,key):
    '''
    id              :ハンドル
    key             :送るキーの配列（最初だけ押しっぱなしで実施する）

    '''
    import win32api
    import win32con

    ctrlwin(id,"fs_active")
    win32api.keybd_event(key[0], 0, 0, 0)#キーダウン
    sleep (50)

    for i in range(1,len(key)):
        win32api.keybd_event(key[i], 0, 0, 0)#キーダウン
        win32api.keybd_event(key[i], 0, win32con.KEYEVENTF_KEYUP, 0)#キーアップ
        sleep (50)

    win32api.keybd_event(key[0], 0, win32con.KEYEVENTF_KEYUP, 0)#キーアップ


def kbd(id,key):
    '''
    id          :ハンドル
    key         :送るキーの配列
    '''

    import win32api
    import win32con

    ctrlwin(id,"fs_active")

    for i in range(0,len(key)):
        win32api.keybd_event(key[i], 0, 0, 0)#キーダウン
        win32api.keybd_event(key[i], 0, win32con.KEYEVENTF_KEYUP, 0)#キーアップ
        sleep (20)


def kbd2(key):
    '''
    id          :ハンドル
    '''

    import win32api
    import win32con

    for i in range(0,len(key)):
        win32api.keybd_event(key[i], 0, 0, 0)#キーダウン
        win32api.keybd_event(key[i], 0, win32con.KEYEVENTF_KEYUP, 0)#キーアップ
        sleep (20)



def sleep(ms):
    '''
    sleep("ミリ秒")
    '''

    time.sleep(ms/1000)


def getuser():
    '''
    ログインユーザーの取得
    '''
    getdata = getpass.getuser()
    return getdata


def get_computer_name():
    '''
    コンピューター名の取得
    '''
    import socket

    # コンピュータ名を取得
    host = socket.gethostname()
    ##print(host)
    return host


def getdir(path,name):
    '''
    path            :探すパス
    name            :”/”フォルダ検索
                    :ファイルは"*.py"などで指定
    '''

    import glob
    import os

    g.getdir_files={}
    folderfile = os.listdir(path)

    #フォルダの場合
    if name=="/":
        folder = [f for f in folderfile if os.path.isdir(os.path.join(path, f))]
        p=0
        for name in folder:
            ###print(name)
            g.getdir_files[p]=name
            p=p+1

        return p

    #ファイルの場合
    else:
        p=0
        for name in glob.glob(path+"/"+name,recursive=True):
            ###print(name)
            p1=pos("\\",name,-1)
            name3=copy(name,p1+1,len(name))
            g.getdir_files[p]=name3
            p=p+1

        return p


def IsNum(moji):
    '''
    数字かどうか判定
    True            :数字
                    :数字ではない
    '''
    return moji.isdigit()


def IsNull(moji):
    '''
    Nullかどうか判定
    True            :Null
                    :Nullではない
    '''
    if moji==None:
        return True
    else:
        return False


def IsAlpha(moji):
    '''
    アルファベットかどうか判定
    True            :アルファベット
                    :アルファベットではない
    '''
    if moji.isalpha() == True:
        if len(moji) != len(moji.encode('utf-8')):
            return False
        else:
            return True
    else:
        return False


def Isdir(path):
    '''
    フォルダかどうか判定
    True            :フォルダ
                    :フォルダではない
    '''
    ##os.path.isfile(path)
    return os.path.isdir(path)


def IsTyp(hensuu):
    '''
    型を返す
    '''
    getdata=type(hensuu)
    getdata=str(getdata)
    getdata=betweenstr(getdata,"'","'",1)
    return getdata


def hash_data_in(hash_table, hash_key, hash_val):
    '''
    辞書型配列に値を入れる
    hash_table      :配列
    hash_key        :キー
    hash_val        :値
    '''
    hash_table[hash_key]=hash_val
    return hash_table


def hash_data_out(hash_table, kind, num):
    '''
    辞書型配列から取り出す・値があるか確認する
    hash_table  :配列
    kind        :"hash_key"   キー
                :"hash_val"   値
                :"hash_exists"   存在するか
    num         :順列番号またはexists時のkey
    '''

    if kind=="hash_key":
        if type(num)==str:
            return "NG"

        cyc=0
        for k in hash_table.keys():
            if cyc==num:break
            cyc=cyc+1

        return k

    elif kind=="hash_val":
        if type(num)==str:
            return "NG"

        cyc=0
        for k in hash_table.values():
            if cyc==num:break
            cyc=cyc+1

        return k

    elif kind=="hash_exists":
        return num in hash_table


def hash_data_del(hash_table,hash_atai=""):
    '''
    hash_table          : 配列
    hash_atai           :"削除するキー
                        :""で全削除
    '''

    if hash_atai=="":
        hash_table.clear()
        return len(hash_table)

    else:
        del hash_table[hash_atai]
        return len(hash_table)


def hash_sort(hashtable,kind=""):
    '''
    hashtable       :配列
    kind            :""             昇順
                    :"reverse"      降順
    '''
    if kind=="":
        hashtable2 = dict(sorted(hashtable.items()))
        return hashtable2
    elif kind=="reverse":
        hashtable2 = dict(sorted(hashtable.items(),reverse=True))
        return hashtable2


def fopen(file,kind,kind2=0):
    '''
    f1={}
    f1=fopen(path,"f_read_or_f_write")

    file            :ファイルのパス
    kind            :"f_exists";存在有無
                    :"f_read":読む
                    :"f_write":新規
                    :"f_read_or_f_write":読み書き
    kind2           :0   UTF8
                    :1   shift_jis
    '''

    fdata={}

    if kind=="f_exists":
        if os.path.exists(file) == False:
            return False
        else:
            return True
    if kind=="f_read" or kind=="f_read_or_f_write":
        if os.path.exists(file) == False:
            fdata["0_0"]=file
            fdata["0_1"]=0
            fdata["0_2"]=0
            fdata["0_3"]=""
            return fdata
        else:
            if kind2 == 0:
                f = open(file, 'r', encoding='utf-8-sig')
            elif kind2==1:
                f = open(file, 'r', encoding='shift_jis')

            gyou=1
            max_retu=1

            for line in f:
                line = line.strip() #前後空白削除
                line = line.replace('\n','') #末尾の\nの削除
                line = line.split(",") #分割

                retu=1
                for col in range(0,len(line)):
                    getdata=line[col]

                    fdata[str(gyou)+"_"+str(retu)]=getdata
                    retu=retu+1

                gyou=gyou+1
                if max_retu < (retu-1):
                    max_retu=(retu-1)

            f.close()
            fdata["0_0"]=file
            fdata["0_1"]=(gyou-1)
            fdata["0_2"]=max_retu
            if kind=="f_read":
                fdata["0_3"]="read_only"
            else:
                fdata["0_3"]=""

            return fdata
    if kind=="f_write":
        fdata["0_0"]=file
        fdata["0_1"]=0
        fdata["0_2"]=0
        fdata["0_3"]=""
        return fdata


def fget(fdata,gyou,retu=1):
    '''
    fget(f1,1,1)

    fdata           :fopenで開いたファイル配列
    gyou            :ファイルの行
    retu            :ファイルの列
    '''

    if gyou==-1:
        return fdata["0_1"]

    elif gyou==-2:

        row=fdata["0_1"]
        col=fdata["0_2"]

        f_str=""
        for i in range(1,row+1):
            for ii in range(1,col+1):
                if ii==1:
                    try:
                        f_str=f_str+fdata[str(i)+"_"+str(ii)]
                    except KeyError:
                        fdata[str(i)+"_"+str(ii)]=""
                        f_str=f_str+fdata[str(i)+"_"+str(ii)]
                else:
                    try:
                        f_str=f_str+','+fdata[str(i)+"_"+str(ii)]
                    except KeyError:
                        fdata[str(i)+"_"+str(ii)]=""
                        f_str=f_str+fdata[str(i)+"_"+str(ii)]
            if i==row:
                f_str=f_str
            else:
                f_str=f_str+'\n'

        return f_str
    else:
        try:
            getdata=fdata[str(gyou)+"_"+str(retu)]
        except:
            fdata[str(gyou)+"_"+str(retu)]=""

        return fdata[str(gyou)+"_"+str(retu)]


def fput(fdata,atai,gyou=0,retu=0):
    '''
    fput(f1,atai,1,1)

    fdata           :fopenで開いたファイル配列
    atai            :値
    gyou            :ファイルの行
    retu            :ファイルの列
    '''

    if fdata["0_3"]=="read_only":
        msg("読み取り専用です",0)
    elif gyou != 0 and retu != 0:
        fdata[str(gyou)+"_"+str(retu)]=atai

        e_gyou=fdata["0_1"]
        e_retu=fdata["0_2"]

        if e_gyou<gyou:
            e_gyou=gyou
            fdata["0_1"]=e_gyou

        if e_retu<retu:
            e_retu=retu
            fdata["0_2"]=e_retu

    else:
        gyou=1
        max_retu=1


        atai = atai.strip() #前後空白削除
        row_str = atai.split('\n') #分割

        gyou=1
        max_retu=1

        for i in range(0,len(row_str)):
            retu=1
            col_str=row_str[i].split(',')

            for ii in range(0,len(col_str)):
                fdata[str(gyou)+"_"+str(retu)]=col_str[ii]
                retu=retu+1

            gyou=gyou+1
            if max_retu < (retu-1):
                max_retu=(retu-1)

        fdata["0_1"]=(gyou-1)
        fdata["0_2"]=max_retu


def fclose(fdata,kind2=0):
    '''
    fclose(f1)

    fdata           :fopenで開いたファイル配列
    kind2           :0      UTF8
                    :1      SHIFT-JIS
    '''
    if fdata["0_3"]=="read_only":
        pass
    else:
        path=fdata["0_0"]

        row=fdata["0_1"]
        col=fdata["0_2"]


        f_str=""
        for i in range(1,row+1):
            for ii in range(1,col+1):
                if ii==1:
                    try:
                        f_str=f_str+str(fdata[str(i)+"_"+str(ii)])
                    except KeyError:
                        fdata[str(i)+"_"+str(ii)]=""
                        f_str=f_str+str(fdata[str(i)+"_"+str(ii)])
                else:
                    try:
                        f_str=f_str+','+str(fdata[str(i)+"_"+str(ii)])
                    except KeyError:
                        fdata[str(i)+"_"+str(ii)]=""
                        f_str=f_str+','+str(fdata[str(i)+"_"+str(ii)])

            f_str=f_str+'\n'

        if kind2 == 0:
            f = open(path, 'w', encoding='UTF-8')
        elif kind2==1:
            f = open(path, 'w', encoding='shift_jis')
        f.writelines(f_str)
        f.close()


def fdelline(fdata,gyou):
    '''
    fdata           :fopenで開いたファイル配列
    gyou            :削除する行
    '''

    path=fdata["0_0"]
    row=fdata["0_1"]
    col=fdata["0_2"]

    fdata2 = {}

    t=1
    for i in range(1,row+1):
        if gyou == i:
            continue
        for ii in range(1,col+1):
                fdata2[str(t)+"_"+str(ii)]=fdata[str(i)+"_"+str(ii)]

        t=t+1

    fdata.clear()

    fdata["0_0"]=path
    fdata["0_1"]=(t-1)
    fdata["0_2"]=col

    for i in range(1,t):
        for ii in range(1,col+1):
            try:
                fdata[str(i)+"_"+str(ii)]=fdata2[str(i)+"_"+str(ii)]
            except KeyError:
                fdata[str(i)+"_"+str(ii)]=""
                fdata[str(i)+"_"+str(ii)]=fdata2[str(i)+"_"+str(ii)]


def get_gamen_size():
    '''
    画面サイズを得る
    '''
    import pyautogui
    screen_w,screen_h = pyautogui.size()
    g.g_screen_w=screen_w
    g.g_screen_h=screen_h


def str_conv(moji,kind):
    '''
    moji        :変換する文字
    kind        :0    大文字⇒小文字
                :1    小文字⇒大文字
                :2    全角⇒半角
                :3    半角⇒全角
    '''
    import  jaconv

    if kind==0:
        return moji.lower()
    elif kind==1:
        return moji.upper()
    elif kind==2:
        return jaconv.z2h(moji,kana=True, digit=True, ascii=True)
    elif kind==3:
        return jaconv.h2z(moji,kana=True, digit=True, ascii=True)


def copy(moji,num1,num2):
    '''
    moji                :コピー元の文字列
    num1                :開始文字位置
    num2                :コピー文字数
    '''

    num2=num1+num2
    return moji[(num1-1):(num2-1)]


def replace(moji,str1,str2):
    '''
    moji            :元の文字列
    str1            :検索する文字
    str2            :置換する文字
    '''

    return moji.replace(str1,str2)


def token(kugiri,moji):
    '''
    kugiri          :区切り文字
    moji            :元の文字=======g.token_strに入れておくこと
    '''


    token_str=moji
    pos=token_str.find(kugiri)

    if pos==-1:
        moji2=token_str
        g.token_str=""
    else:
        moji2=token_str[0:pos]
        moji3=token_str[(pos+1):len(token_str)]

        while True:
            if moji3[0:1]==kugiri:
                moji3=moji3[1:len(moji3)]
            else:
                break

        g.token_str=moji3

    return moji2


def pos(s_moji,moji,num=1):
    '''
    s_moji              :見つける文字
    moji                :探す文字
    num                 :何個目か   マイナスは後ろから
    '''

    pos1=0

    if num > 0:
        for i in range(num):
            pos1=moji.find(s_moji,pos1,len(moji))

            if pos1==-1:
                break

            if i != (num-1):
                 pos1=pos1+1


    else:           #後ろから検索
        num=abs(num)
        for i in range(num):
            pos1=moji.rfind(s_moji)

            if pos1==-1:
                break

            moji1 = moji[0:pos1]
            moji=moji1

    return pos1+1


def betweenstr(strA,mojiA,mojiB,num):
    '''
'    betweenstr(str, mojiA, mojiB, num) As String
'    str            :探す元になる文字列
'    mojiA          :得たい文字列の前にある文字列
'    mojiB          :得たい文字列の後にある文字列
'    num            :n個目の該当文字列を返す (マイナス値で指定すると後ろからサーチ)

'    戻り値；取り出した文字　ない場合は　"false"
    '''

    if num > 0 :
        for i in range(0,num):
            pos1 = pos(mojiA,strA,1)

            if pos1 == 0 :
                return "false"

            pos1 = pos1 + len(mojiA)
            strB=copy(strA,pos1,len(strA))

            pos2=pos(mojiB,strB,1)
            if pos2 == 0:
                return "false"

            pos2=pos1+pos2-1

            betweenstr = copy(strA, pos1, pos2-pos1)
            strA = copy(strA, pos2+1, len(strA))
        return betweenstr
    elif num < 0:
        for i in range(0,abs(num)):
            pos1 = pos(mojiA,strA,-1)
            pos2 = pos(mojiB,strA,-1)

            if pos1 == 0 :
                return "false"

            if pos1==pos2:
                strB=copy(strA,1,pos1-1)
                pos3=pos(mojiA,strB,-1)

                if pos3 == 0 :
                    return "false"

                betweenstr=copy(strA,pos3+1,(pos2-pos3-1))
                strA=copy(strA,1,pos3-1)
            else:
                betweenstr=copy(strA,pos1+1,(pos2-pos1-1))
                strA=copy(strA,1,pos1-1)
        return betweenstr
    else:
        return "false"


def File_system(kind,pathA,pathB=""):
    '''
    kind
        "fs_createfolder"           :フォルダ作成
        "fs_createtextfile"         :ファイル作成
        "fs_copyfile"               :ファイル作成
        "fs_copyfolder"             :ファイル作成
        "fs_move"                   :移動
        "fs_exists"                 :存在確認
        "fs_deletefile"             :ファイル削除
        "fs_deletefolder"           :フォルダ削除
        "fs_create_time"            :作成日
        "fs_update_time"               :更新日
        "fs_access_time"               :アクセス日
        "fs_file_size"              :ファイルサイズ
        "fs_folder_size"            :フォルダサイズ
        "fs_readonly_check"         :読み取り専用かどうか　 　True：読み取り専用
        "fs_readonly_kill"          :読み取り専用を解除
        "fs_readonly_do"            :読み取り専用にする
    '''
    import shutil
    import os
    from pathlib import Path
    import datetime
    import stat

    if kind=="fs_createfolder":
        os.makedirs(pathA)
    elif kind=="fs_createtextfile":
        f = open(pathA, 'w')
        f.write('')  # 何も書き込まなくてファイルは作成されました
        f.close()
        return True
    elif kind=="fs_copyfile":
        if pathB=="":
            msg("コピー先を指定してください",0)
            return False
        else:
            shutil.copy2(pathA,pathB)
            return True
    elif kind=="fs_move":
        shutil.move(pathA,pathB)
    elif kind=="fs_copyfolder":
        if pathB=="":
            msg("コピー先を指定してください",0)
            return False
        else:
            shutil.copytree(pathA,pathB)
            return True
    elif kind=="fs_exists":
        if os.path.exists(pathA)==False:
            return False
        else:
            return True
    elif kind=="fs_deletefile":
        os.remove(pathA)
        return True
    elif kind=="fs_deletefolder":
        shutil.rmtree(pathA)
        return True
    elif kind=="fs_create_time":
        p = Path(pathA)
        create_time = datetime.datetime.fromtimestamp(p.stat().st_ctime)
        return create_time
    elif kind=="fs_update_time":
        p = Path(pathA)
        update_time = datetime.datetime.fromtimestamp(p.stat().st_mtime)
        return update_time
    elif kind=="fs_access_time":
        p = Path(pathA)
        access_time = datetime.datetime.fromtimestamp(p.stat().st_atime)
        return access_time
    elif kind=="fs_file_size":
        file_size = os.path.getsize(pathA)
        return file_size
    elif kind=="fs_folder_size":
        total_size = 0
        for dir_path in os.listdir(pathA):
            full_path = os.path.join(pathA, dir_path)
            if os.path.isfile(full_path):
                total_size += os.path.getsize(full_path)
            elif os.path.isdir(full_path):
                total_size += get_size_dir(full_path)
        return total_size
    elif kind=="fs_readonly_check":
        return not os.access(pathA, os.W_OK)
    elif kind=="fs_readonly_kill":
        os.chmod(pathA, stat.S_IWRITE)
        return True
    elif kind=="fs_readonly_do":
        os.chmod(pathA, stat.S_IREAD)
        return True


def msg(msg_str,msg_title="msg",kind=0):
    '''
    msg_title       :タイトル
    msg_str         :メッセージ
    kind            :0      情報
                    :1      警告
                    :2      エラー
                    :3      OKボタンのみ
                    :4      はい・いいえ              はいボタン・・・"yes"   いいえボタン・・・"no"
                    :5      OK・キャンセル            OKボタン・・・True   キャンセルボタン・・・False
                    :6      再試行・キャンセル）      再試行ボタン・・・True  キャンセルボタン・・・False
    '''

    import tkinter.messagebox as messagebox
    tk.Tk().withdraw()

    th1=Thread(msg_top,msg_title)

    # メッセージボックス（情報）
    if kind==0:
        ret=messagebox.showinfo(msg_title,msg_str)

    # メッセージボックス（警告）
    elif kind==1:
        ret=messagebox.showwarning(msg_title,msg_str)

    # メッセージボックス（エラー）
    elif kind==2:
        ret=messagebox.showerror(msg_title,msg_str)

    # メッセージボックス（はい・いいえ）
    elif kind==3:
        ret=messagebox.askyesno(msg_title,msg_str)

    # メッセージボックス（はい・いいえ）
    elif kind==4:
        ret=messagebox.askquestion(msg_title,msg_str)

    # メッセージボックス（OK・キャンセル）
    elif kind==5:
        ret=messagebox.askokcancel(msg_title,msg_str)

    # メッセージボックス（再試行・キャンセル）
    elif kind==6:
        ret=messagebox.askretrycancel(msg_title,msg_str)

    Thread_end(th1)

    return ret


def msg_top(h1):
    '''
    msgボックスを最前面に出すためのスレッド用def
    '''
    while True:
        hwnd=getid(h1)
        if hwnd > 0:break

    ctrlwin(hwnd,"fs_topmost")


def inputbox(msg_str,msg_title="Message",kind=0,default_str=""):
    '''
    num = j.inputbox("何行目ですか？")

    msg_str             :メッセージ
    msg_title           :タイトル
    kind                :0  普通のINPUTBOX
                        :0以外　　  ファイル選択ダイアログ
    default_str         :デフォルト文字
    戻り値               :値  もしくは　　フルパス
                        :空欄時 ""
                        :キャンセル時 None
    '''

    import tkinter.simpledialog as simpledialog
    import tkinter.filedialog as filedialog

    tk.Tk().withdraw()

    # 入力ダイアログ
    if kind==0:
        ret = simpledialog.askstring(msg_title,msg_str,initialvalue=default_str)
    elif kind==1:
        ret = filedialog.askopenfilename()

    return ret


def password(msg_str,default_str="",msg_title="Password"):
    '''
    getdata=password("パスワードを入力してください","Jinjin","Password")
    msg_str             :メッセージ分
    default_str         :デフォルト入力文字
    msg_title           :タイトル
    '''

    import PySimpleGUI as sg

    menu_def = None
    ldict = {}

    sg.theme('Default1')
    menu_def = "g.menu_def='"+msg_str+"'"
    exec(menu_def, globals(), ldict)


    layout = [
    [sg.Text(g.menu_def,font=('メイリオ',12),size=(25, 1)),sg.Input(default_str, size=(15, 1), password_char="*", key="-password-")],
    [sg.Checkbox("パスワードを表示します", enable_events=True, key="-toggle_password-")],
    [sg.Button("OK", size=(5, 1), enable_events=True, key="-OK-")]
    ]

    gw=int(g.g_screen_w*0.3)
    window = sg.Window(msg_title,layout,resizable=True,size=(gw, 100), location=(0, 0))

    while True:  ## Event Loop
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        elif event == "-toggle_password-":
            # チェックボタンがONだったらパスワードを表示する
            if values["-toggle_password-"]:
                window["-password-"].update(password_char="")

            # チェックボタンがOFFだったらパスワードを非表示にする
            else:
                window["-password-"].update(password_char="*")
        elif event == "-OK-":
            password=values["-password-"]
            break


    window.close()
    return password


def gettime(g_day,kijun=""):
    '''
    g_day           :0      今日
                    :1      1日後
                    :-1     1日前
    kijun           :"20230401"
    '''

    import datetime

    if kijun=="":
        kijun=str(datetime.datetime.now())
##        print(kijun)
        kijun_nen=int(copy(kijun,1,4))
        kijun_tuki=int(copy(kijun,6,2))
        kijun_nichi=int(copy(kijun,9,2))
    else:
        kijun_nen=int(copy(kijun,1,4))
        kijun_tuki=int(copy(kijun,5,2))
        kijun_nichi=int(copy(kijun,7,2))

    day1=datetime.datetime.now()
##    print(day1)
    day2=datetime.datetime(kijun_nen,kijun_tuki,kijun_nichi)
##    print(day2)
    days=day1-day2
    days=str(days)
##    print(days)

    if pos("day",days)==0:
        days=0
    elif pos(" days",days)==0:
        pos1=pos(",",days,1)
        days=copy(days,1,pos1-1)
        days=int(replace(days," day",""))
    else:
        pos1=pos(",",days,1)
        days=copy(days,1,pos1-1)
        days=int(replace(days," days",""))

    g_day = g_day*-1
    g_day+=days

    td=str(datetime.datetime.now()-datetime.timedelta(days=g_day))
    g.token_str=td
    g.G_TIME_YY4=token("-",g.token_str)
    g.G_TIME_YY2=g.G_TIME_YY4[2:4]
    g.G_TIME_MM2=token("-",g.token_str)
    g.G_TIME_DD2=token(" ",g.token_str)
    g.G_TIME_HH2=token(":",g.token_str)
    g.G_TIME_NN2=token(":",g.token_str)
    g.G_TIME_SS2=token(".",g.token_str)
    g.G_TIME_YYYY=int(g.G_TIME_YY4)
    g.G_TIME_YY=int(g.G_TIME_YY2)
    g.G_TIME_MM=int(g.G_TIME_MM2)
    g.G_TIME_DD=int(g.G_TIME_DD2)
    g.G_TIME_HH=int(g.G_TIME_HH2)
    g.G_TIME_NN=int(g.G_TIME_NN2)
    g.G_TIME_SS=int(g.G_TIME_SS2)

    date = datetime.date(g.G_TIME_YYYY,g.G_TIME_MM,g.G_TIME_DD)
    g.G_TIME_WW = date.strftime('%a')


def gettime2():
    '''
    起動からの経過時間(ms)
    '''
    time1=time.perf_counter()
    time1*=1000
    return time1


def fukidasi(msg="",color="skyblue"):
    '''
    fukidasi(driver0,msg,color)

    driver0     :driver0
    msg         :表示する文字
                :省略でfukidasiを消去
    color       :背景色
                white,lightyellow,skyblue,lightgreen,pink,silver,blue,aqua,yellow,orange,red,gray,fuchsia,brownなど
    '''

    import warnings

    warnings.filterwarnings('ignore')
    if msg != "":

        IE_msg_str = "<html>"
        IE_msg_str=IE_msg_str+"<head>"
        IE_msg_str=IE_msg_str+"<title> pythonFUKI </title>"
        IE_msg_str=IE_msg_str+"<style type='text/css'>"
        IE_msg_str=IE_msg_str+"body{background-color:"+color+";}"
        IE_msg_str=IE_msg_str+"</style>"
        IE_msg_str=IE_msg_str+"</head>"
        IE_msg_str=IE_msg_str+"<body>"
        IE_msg_str=IE_msg_str+"<span id='msg'>"+msg+"</span></body></html>"

        f1={}
        f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
        fput(f1, IE_msg_str)
        fclose(f1)

        ###########IE_msg###########
        try:
            g.driver0.set_window_size(g.g_screen_w, 200)
            g.driver0.get(g.CUR_DIR+"//Temp.html")
            BusyWait(g.driver0)
        except:
            g.driver0 = webdriver.Edge(service=g.service, options=options)
            g.driver0.set_window_size(g.g_screen_w, 200)
            g.driver0.get(g.CUR_DIR+"//Temp.html")
            BusyWait(g.driver0)
        ###########IE_msg###########

        win=getid("all_win")
        for i in range(0,len(g.all_win)):
            if pos("pythonFUKI",g.all_win[i])>0:
                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                break

        ctrlwin(id,"fs_show")
        ACW(id, 0, 0, g.g_screen_w, 200)

    else:
        win=getid("all_win")
        for i in range(0,len(g.all_win)):
            if pos("pythonFUKI",g.all_win[i])>0:
                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                break
        ctrlwin(id,"fs_min")
        g.driver0.execute_script("document.getElementById('msg').innerText=''")
        ##g.driver_fuki.close()
        ##g.driver_fuki.quit()

    warnings.filterwarnings('default')


def fukidasi2(msg=""):
    '''
    msg             :メッセージ
    '''
    import PySimpleGUI as sg

    def fuki2():
        menu_def = None
        ldict = {}

        sg.theme('Default1')
        menu_def = "g.menu_def='"+msg+"'"
        exec(menu_def, globals(), ldict)


        layout = [
        [sg.Text(g.menu_def,font=('メイリオ',12),size=(50, 1))]
        ]

        gw=int(g.g_screen_w*0.9)
        window = sg.Window('Theme Browser',layout,resizable=True,size=(gw, 100), location=(0, 0))

        while True:  ## Event Loop
            event, values = window.read()

            if event in (sg.WIN_CLOSED, 'Exit'):
                break
            if g.fuki_kill==True:
                break

        window.close()


    if msg == "":
        g.fuki_kill=True
        ctrlwin(getid("Theme Browser"),"fs_close")
        Thread_end(g.th100)
        return ""
    else:
        g.th100=Thread(fuki2)


def IE_msg(msg,btn_type="btn_ok",sizex=500,sizey=150,TimeOut=600):
    '''
    IE_msg(msg,[btn_type,sizex,sizey,timeout])

    msg             :文字
    btn_type        :"btn_ok"
                    :"btn_ok_or_btn_no"
    sizex           :横幅
    sizey           :縦幅
    timeout         :タイムアウトまでの時間

    戻り値；"btn_ok" 又は　"btn_no" 又は　"timeout"
    '''



    IE_msg_str="<html>\n"
    IE_msg_str=IE_msg_str+"<head>\n"
    IE_msg_str=IE_msg_str+"<title>選択してください。</title>\n"
    IE_msg_str=IE_msg_str+"<script language='JavaScript'>\n"
    IE_msg_str=IE_msg_str+"function event(num){document.getElementById('event').innerText=num}\n"
    IE_msg_str=IE_msg_str+"</script>\n"
    IE_msg_str=IE_msg_str+"<style type='text/css'><!--\n"
    IE_msg_str=IE_msg_str+"table{border-style:solid;border-width:1px;border-color:black;border-collapse: collapse;font-size:12;Text -Align: center; padding:0;}\n"
    IE_msg_str=IE_msg_str+"th{border-style:solid;border-width:1px;border-color:black;padding:0;background-color:#ccccff;text-align:center;height:20;}\n"
    IE_msg_str=IE_msg_str+"td{border-style:solid;border-width:1px;border-color:black;padding:0;background-color:#ffffcc;text-align:center;height:12;}\n"
    IE_msg_str=IE_msg_str+"--></style></head>\n"
    IE_msg_str=IE_msg_str+"<body width='30' height='60'>\n"
    IE_msg_str=IE_msg_str+"<span id='alete'><font size='4'>"+msg+"</font></span>"

    if btn_type=="btn_ok":
        IE_msg_str = IE_msg_str+"　　<br><br><input id='btn1' type='submit' value = 'OK' onClick='value="+chr(34)+"-"+chr(34)+"'>"
    elif btn_type=="btn_ok_or_btn_no":
        IE_msg_str = IE_msg_str+"　　<br><br><input id='btn1' type='submit' value = 'OK' onClick='value="+chr(34)+"-"+chr(34)+"'>"
        IE_msg_str = IE_msg_str+"　　<input id='btn2' type='button' value = 'NO' onClick='value="+chr(34)+"-"+chr(34)+"'>"

    IE_msg_str = IE_msg_str+"</body>"
    IE_msg_str = IE_msg_str+"</html>"

    f1={}
    f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
    fput(f1, IE_msg_str)
    fclose(f1)


    ###########IE_msg###########
    try:
        g.driver0.set_window_size(sizex + 400, sizey + 100)
        g.driver0.get(g.CUR_DIR+"/Temp.html")
        BusyWait(g.driver0)

        win=getid("all_win")
        for i in range(0,len(g.all_win)):
            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                break
        ctrlwin(id,"fs_show")
    except:
        g.driver0 = webdriver.Chrome(service=g.service, options=options)
        win=getid("all_win")
        for i in range(0,len(g.all_win)):
            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                break
        ctrlwin(id,"fs_show")
        g.driver0.set_window_size(sizex + 400, sizey + 100)
        g.driver0.get(g.CUR_DIR+"/Temp.html")
        BusyWait(g.driver0)
    ###########IE_msg###########



    getdata=g.driver0.execute_script("return document.getElementById('btn1').focus()")


    time.sleep(100/1000)

    try:
        msg_cyc = 0
        while True:
            getdata = g.driver0.execute_script("return document.getElementById('btn1').value")

            if getdata == "-":
                win=getid("all_win")
                for i in range(0,len(g.all_win)):
                    if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                        id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                        break
                ctrlwin(id,"fs_min")
                g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                return "btn_ok"

            if btn_type == "btn_ok_or_btn_no":
                getdata = g.driver0.execute_script("return document.getElementById('btn2').value")

                if getdata == "-":
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    return "btn_no"

            sleep (100)

            if ((msg_cyc / 10) > TimeOut):
                ##g.driver0.quit()
                win=getid("all_win")
                for i in range(0,len(g.all_win)):
                    if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                        id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                        break
                ctrlwin(id,"fs_min")
                g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                return "timeout"

            msg_cyc = msg_cyc + 1



    except:
        ##g.driver0.quit()
        return "IE_close"


def IE_input(msg,default="",path_sen=False,password=False):
    '''
    IE_input(msg,default,[path_sen,pasword])

        msg                    :文字
        default                :デフォルトの入力文字
        path_sen               :false；通常のinputbox
                               :true；ディレクトリ選択画面にする
        pasword                :true；入力文字を*で表示
                               :false；入力文字は通常表示

        戻り値；入力された文字　キャンセル時はfalse
    '''



    TimeOut = 600

    #通常のinputbox
    if path_sen == False:

        IE_msg_str = "<html>\n"
        IE_msg_str=IE_msg_str+"<head>\n"
        IE_msg_str=IE_msg_str+"<title>入力してください。</title>\n"
        IE_msg_str=IE_msg_str+"<script language='JavaScript'>\n"
        IE_msg_str=IE_msg_str+"function eventA(){getdata=document.getElementById('atai').value\n"
        IE_msg_str=IE_msg_str+"document.getElementById('eventAA').innerText=getdata\n"
        IE_msg_str=IE_msg_str+"}\n"
        IE_msg_str=IE_msg_str+"</script>\n"
        IE_msg_str=IE_msg_str+"</head>\n"
        IE_msg_str=IE_msg_str+"<body width='30' height='60'>\n"
        IE_msg_str=IE_msg_str+"<span id='alete'><font size='4'>"+msg+"</font></span>"

        if password == False:
            IE_msg_str = IE_msg_str+"<br><br><input id='atai' type='text' onchange='eventA()' value = '"+default+"'>\n"
        else:
            IE_msg_str = IE_msg_str+"<br><br><input id='atai' type='password' onchange='eventA()' value = '"+default+"'>\n"

        IE_msg_str = IE_msg_str+"<br><br><br><input id='btn1' type='button' value = 'ok' onClick='value="+chr(34)+"-"+chr(34)+ "'>\n"
        IE_msg_str = IE_msg_str+"　　　<input id='btn2' type='button' value = 'cancel' onClick='value=" +chr(34)+"-"+chr(34)+"'>\n"

        IE_msg_str = IE_msg_str+"<span id='eventAA'  style='display:none;'></span></body>\n"
        IE_msg_str = IE_msg_str+"</html>"



        f1={}
        f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
        fput(f1, IE_msg_str)
        fclose(f1)


        ###########IE_msg###########
        try:
            g.driver0.set_window_size(g.g_screen_w * 4 / 5, 400)
        except:
            g.driver0 = webdriver.Edge(service=g.service, options=options)
            g.driver0.set_window_size(g.g_screen_w * 4 / 5, 400)
        ###########IE_msg###########


        g.driver0.get(g.CUR_DIR+"//Temp.html")
        BusyWait(g.driver0)

        g.driver0.execute_script("document.getElementById('atai').focus()")

        sleep (100)

        if default != "":
            g.driver0.execute_script("document.getElementById('atai').Select")

        try:
            input_cyc = 0
            while True:
                getdata = g.driver0.execute_script("return document.getElementById('btn1').value")
                if getdata == "-":
                    getdata = g.driver0.execute_script("return document.getElementById('atai').value")
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    return getdata


                getdata = g.driver0.execute_script("return document.getElementById('btn2').value")

                if getdata == "-":
                    ##g.driver0.quit
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return False

                sleep (100)

                if (input_cyc / 10) > TimeOut:
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    return "timeout"

                getdata = g.driver0.execute_script("return document.getElementById('eventAA').innerText")
                if getdata != "":
                    ##g.driver0.quit
                    getdata = g.driver0.execute_script("return document.getElementById('atai').value")
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    return getdata

                input_cyc = input_cyc + 1


        except:
            return "web_close"



    #ディレクトリ選択画面にする
    else:

        IE_msg_str="<!DOCTYPE html><html>\n"
        IE_msg_str=IE_msg_str+"<head>\n"
        IE_msg_str=IE_msg_str+"<title>入力してください。</title>\n"
        IE_msg_str=IE_msg_str+"<script language='JavaScript'>\n"
        IE_msg_str=IE_msg_str+"function eventA(){getdata=document.getElementById('atai').value\n"
        IE_msg_str=IE_msg_str+"document.getElementById('eventAA').innerText=getdata\n"
        IE_msg_str=IE_msg_str+"}\n"
        IE_msg_str=IE_msg_str+"</script>\n"
        IE_msg_str=IE_msg_str+"</head>\n"
        IE_msg_str=IE_msg_str+"<body width='30' height='60'>\n"
        IE_msg_str=IE_msg_str+"<span id='alete'><font size='4'>"+msg+"</font></span>\n"

        IE_msg_str=IE_msg_str+"<br><br><input id='atai' onchange='eventA()' type='file' name='upfile[]' accept='text/plain'  value = '"+default+"' multiple='multiple'>"

        IE_msg_str=IE_msg_str+"<br><br><br><input id='btn1' type='submit' value = 'ok' onClick='value="+chr(34)+"-"+chr(34)+ "'>"
        IE_msg_str=IE_msg_str+"　　　<input id='btn2' type='button' value = 'cancel' onClick='value="+chr(34)+"-"+chr(34)+"'>"

        IE_msg_str=IE_msg_str+"<span id='eventAA'></span></body>"
        IE_msg_str=IE_msg_str+"</html>"



        f1={}
        f1 = fopen(g.CUR_DIR+"//\Temp.html","f_write")
        fput(f1,IE_msg_str)
        fclose(f1)

        ############IE_msg###########
        try:
            g.driver0.set_window_size(g.g_screen_w * 4/5, 400)
        except:
            g.driver0 = webdriver.Edge(service=g.service, options=options)
            g.driver0.set_window_size(g.g_screen_w * 4/5, 400)
        ############IE_msg###########


        g.driver0.get(g.CUR_DIR+"//Temp.html")
        BusyWait(g.driver0)


        g.driver0.execute_script("return document.getElementById('btn1').focus()")

        sleep (100)

        try:
            input_cyc = 0
            while True:
                getdata = g.driver0.execute_script("return document.getElementById('btn1').value")

                if getdata == "-":
                    ##g.driver0.quit()
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return g.driver0.execute_script("return document.getElementById('atai').value")

                getdata = g.driver0.execute_script("return document.getElementById('btn2').value")

                if getdata == "-":
                    ##g.driver0.quit()
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return False

                sleep (100)

                if (input_cyc/10) > TimeOut:
                    ##g.driver0.quit()
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').innerText=''")
                    return "timeout"

                getdata = g.driver0.execute_script("return document.getElementById('eventAA').innerText")
                if getdata != "":
                    ##g.driver0.quit
                    getdata = g.driver0.execute_script("return document.getElementById('atai').value")
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("入力してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    g.driver0.execute_script("document.getElementsByTagName('body').item(0).innerText=''")
                    return getdata

                input_cyc = input_cyc + 1

        except:
            return "web_close"


def IE_gettable(driver0,table_num=0,row_num=-1,cell_num=-1):
    '''
    IE_gettable(driver0,table_num,[row_num,cell_num])

    driver0                 :driverのオブジェクト
    table_num               :tableのitem番号
    row_num                 :行
    cell_num                :列

    戻り値                   :取得した文字
                            :行省略時は　行数_列数 を返す

    行数；g.IE_table_rows
    列数；g.IE_table_cells
    データ；g.IE_table(行数 & "_" & 列数)　に格納
    '''

    g.IE_table ={}

    driver0.execute_script("IE_tb=document.getElementsByTagName('table').item("+str(table_num)+")")
    ie_gyou = driver0.execute_script("return IE_tb.rows.length")
    ie_retu = driver0.execute_script("return IE_tb.rows[0].cells.length")

    g.IE_table_rows = ie_gyou
    g.IE_table_cells = ie_retu


    #テーブルの値を配列に入れる
    for i in range(0,ie_gyou):
        for ii in range(0,ie_retu):
            getdata = driver0.execute_script("return IE_tb.rows["+str(i)+"].cells["+str(ii)+"].innerText")
            g.IE_table[str(i)+"_"+str(ii)]= getdata

    #行省略時
    if row_num==-1:
        return str(ie_gyou)+"_"+str(ie_retu)
    else:
        return g.IE_table[str(row_num)+"_"+str(cell_num)]


def SLCTBOX(slct_menu,kind=0,IE_left=0,IE_top=0,IE_w=300,IE_h=300,TimeOut=60,msg="選択してください",return_kind="str"):
    '''
    SLCTBOX(slct_menu,[kind,IE_left,IE_top,IE_w,IE_h,timeout,msg,return_kind])

    slct_menu               :選択のための配列
    kind                    :0      ボタン
                            :1      チェックボックス
                            :2      リストボックス

    IE_left                 :表示位置
    IE_top                  :表示位置
    IE_w                    :表示幅
    IE_h                    :表示高さ
    timeout                 :タイムアウトの時間　省略で60ｓ
    msg                     :表示文字
    return_kind             :戻り値を選択
                            :"str" 項目名（hash_val）を返す
                            :"num" 配列番号を返す
    戻り値；
        選ばれた項目名　　チェックボックス選択時はtabにて結合して返す
       ×閉じ時；-1
       タイムアウトは -2
    '''




    intext = "<html><head>\n"
    intext=intext+"<title>選択してください。</title>\n"
    intext=intext+"<!--<meta http-equiv='x-ua-compatible' content='IE=7' >-->\n"
    intext=intext+"<meta http-equiv='x-ua-compatible' content='IE=EmulateIE7' >\n"
    intext=intext+"<script language='JavaScript'>\n"
    intext=intext+"function eventA(num){document.getElementById('eventA').innerText=num}\n"
    intext=intext+"</script>\n"
    intext=intext+"</head>\n"
    intext=intext+"<body width='30' height='60'>\n"

    ##//###########IE_msg###########
    try:
        g.driver0.set_window_size(IE_w, IE_h)
    except:
        g.driver0 = webdriver.Edge(service=g.service, options=options)
        g.driver0.set_window_size(IE_w, IE_h)
    ##//###########IE_msg###########

    if kind == 0 :
        if type(slct_menu) == dict:
            if return_kind == "str" :
                intext2 = intext + msg + "<br><br>"
                for i in range( 0 , len(slct_menu)):
                    intext2 = intext2 + "<input type='button' id='" + str(i) + "'value='" + hash_data_out(slct_menu, "hash_val", i) + "' onclick = 'eventA(this.value)' style='width:100px;height:20px;margin:5px 0px;'>　　<br>"
            else:
                intext2 = intext + msg + "<br><br>"
                for i in range( 0 , len(slct_menu)):
                    intext2 = intext2 + "<input type='button' id='" + str(i) + "' value='" + hash_data_out(slct_menu, "hash_val", i) + "' onclick = 'eventA(this.id)' style='width:100px;height:20px;margin:5px 0px;'>　　<br>"
        else:
            if return_kind == "str" :
                intext2 = intext + msg + "<br><br>"
                for i in range( 0 , len(slct_menu)):
                    intext2 = intext2 + "<input type='button' id='" + str(i) + "'value='" + slct_menu[i] + "' onclick = 'eventA(this.value)' style='width:100px;height:20px;margin:5px 0px;'>　　<br>"
            else:
                intext2 = intext + msg + "<br><br>"
                for i in range( 0 , len(slct_menu)):
                    intext2 = intext2 + "<input type='button' id='" + str(i) + "' value='" + slct_menu[i] + "' onclick = 'eventA(this.id)' style='width:100px;height:20px;margin:5px 0px;'>　　<br>"

        intext2 = intext2 + "<span id='eventA' style='display:none'></span><br><br>"
        intext2 = intext2 + "<!--<span id='eventA'></span>-->"
        intext2 = intext2 + "</body>"
        intext2 = intext2 + "</html>"

        f1={}
        f1 = fopen(path_join(g.CUR_DIR,"Temp.html"), "f_write")
        fput(f1, intext2)
        fclose(f1)

        g.driver0.get(path_join(g.CUR_DIR,"Temp.html"))
        BusyWait(g.driver0)

        g.driver0.execute_script("return document.getElementById('0').focus()")

        i = 0
        try:
            while True:
                getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")
                if getdata != "":
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return getdata
                if i == 10 * TimeOut :
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return -2 ##タイムアウト
                sleep (100)
                i = i + 1
        except:
            return ""
    elif kind == 1 :
        if type(slct_menu) == dict:
            intext2 = intext + msg + "<br><br>"
            for i in range( 0 , len(slct_menu)):
                intext2 = intext2 + "<input type='checkbox' id='" + hash_data_out(slct_menu, "hash_val", i) + "'>" + hash_data_out(slct_menu, "hash_val", i) + "　　<br>"

            intext2 = intext2 + "<br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'>       <span id='eventA' style='display:none'></span><br><br>"
            intext2 = intext2 + "<!--<span id='eventA'></span>-->"
            intext2 = intext2 + "</body>"
            intext2 = intext2 + "</html>"

            f1={}
            f1 = fopen(path_join(g.CUR_DIR,"Temp.html"),"f_write")

            fput(f1, intext2)
            fclose(f1)

            g.driver0.get(path_join(g.CUR_DIR,"Temp.html"))
            BusyWait(g.driver0)

            i = 0

            try:
                while True:
                    getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")
                    if getdata == "ok" :
                        getdata = ""

                        for t in range(0,len(slct_menu)):
                            if g.driver0.execute_script("return document.getElementById("+chr(34)+hash_data_out(slct_menu, "hash_val", t)+chr(34)+").checked") == True :
                                if return_kind == "str" :
                                    if getdata == "" :
                                        getdata = hash_data_out(slct_menu, "hash_val", t)
                                    else:
                                        getdata = getdata + chr(9) + hash_data_out(slct_menu, "hash_val", t)
                                else:
                                    if getdata == "" :
                                        getdata = str(t)
                                    else:
                                        getdata = getdata + chr(9) + str(t)

                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return getdata

                    if i == 10 * TimeOut :
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return -2 ##タイムアウト

                    sleep (100)

                    i = i + 1
            except:
                return ""
        else:
            try:
                intext2 = intext + msg + "<br><br>"
                for i in range( 0 , len(slct_menu)):
                    intext2 = intext2 + "<input type='checkbox' id='" + slct_menu[i] + "'>" + slct_menu[i] + "　　<br>"

                intext2 = intext2 + "<br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'>       <span id='eventA' style='display:none'></span><br><br>"
                intext2 = intext2 + "<!--<span id='eventA'></span>-->"
                intext2 = intext2 + "</body>"
                intext2 = intext2 + "</html>"

                f1={}
                f1 = fopen(path_join(g.CUR_DIR , "Temp.html"), "f_write")
                fput(f1, intext2)
                fclose(f1)

                g.driver0.get(path_join(g.CUR_DIR,"Temp.html"))
                BusyWait(g.driver0)

                i = 0

                while True:
                    getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")

                    if getdata == "ok" :
                        getdata = ""
                        for t in range( 0 , len(slct_menu)):
                            if g.driver0.execute_script("return document.getElementById("+chr(34)+slct_menu[t]+chr(34)+").checked") == True :
                                if return_kind == "str" :
                                    if getdata == "" :
                                        getdata = slct_menu[t]
                                    else:
                                        getdata = getdata + chr(9) + slct_menu[t]
                                else:
                                    if getdata == "" :
                                        getdata = str(t)
                                    else:
                                        getdata = getdata + chr(9) + str(t)
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return getdata

                    if i == 10 * TimeOut :
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return -2 ##///タイムアウト

                    sleep (100)
                    i = i + 1

            except:
                return ""

    elif kind == 2 :
        msg = msg + "<br>" + "Ctrl で複数選択"

        if type(slct_menu) == dict:
            intext2 = intext + msg + "<br><select size ='" + str(len(slct_menu)) + "' id='SLCTBOX' multiple>"

            for i in range( 0 , len(slct_menu)):
                intext2 = intext2 + "<option value='" + hash_data_out(slct_menu, "hash_val", i) + "'>" + hash_data_out(slct_menu, "hash_val", i) + "</option>"

            intext2 = intext2 + "</select>"
            intext2 = intext2 + "<input type = 'button' value = 'ok' onclick = 'eventA(this.value)' style='position:absolute;left:200;top:10;'>       <span id='eventA' style='display:none'></span><br><br>"
            intext2 = intext2 + "<!--<span id='eventA'></span>-->"
            intext2 = intext2 + "</body>"
            intext2 = intext2 + "</html>"

            f1={}
            f1 = fopen(path_join(g.CUR_DIR , "Temp.html"), "f_write")
            fput(f1, intext2)
            fclose(f1)

            g.driver0.get(path_join(g.CUR_DIR,"Temp.html"))
            BusyWait(g.driver0)

            i = 0
            try:
                while True:
                    getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")

                    if getdata == "ok" :
                        getdata = ""

                        for t in range( 0 , len(slct_menu)):
                            selctbox = g.driver0.execute_script("return document.getElementById('SLCTBOX').options["+str(t)+"].selected")

                            if selctbox == True :
                                if return_kind == "str" :
                                    if getdata == "" :
                                        getdata = hash_data_out(slct_menu, "hash_val", t)
                                    else:
                                        getdata = getdata + chr(9) + hash_data_out(slct_menu, "hash_val", t)
                                else:
                                    if getdata == "" :
                                        getdata = str(t)
                                    else:
                                        getdata = getdata + chr(9) + str(t)
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return getdata

                    if i == 10 * TimeOut :
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return -2 ###///タイムアウト

                    sleep (100)
                    i = i + 1
            except:
                return ""

        else:
            intext2 = intext + msg + "<br><select size ='" + str(len(slct_menu)) + "' id='SLCTBOX' multiple>"

            for i in range( 0 , len(slct_menu)):
                intext2 = intext2 + "<option value='" + slct_menu[i] + "'>" + slct_menu[i] + "</option>"

            intext2 = intext2 + "</select>"
            intext2 = intext2 + "<br><br><input type = 'button' value = 'ok' onclick = 'eventA(this.value)'style='position:absolute;left:200;top:10;'>       <span id='eventA' style='display:none'></span><br><br>"
            intext2 = intext2 + "<!--<span id='eventA'></span>-->"
            intext2 = intext2 + "</body>"
            intext2 = intext2 + "</html>"

            f1={}
            f1 = fopen(path_join(g.CUR_DIR , "Temp.html"), "f_write")
            fput(f1, intext2)
            fclose(f1)

            g.driver0.get(path_join(g.CUR_DIR,"Temp.html"))
            BusyWait(g.driver0)

            i = 0
            try:
                while True:
                    getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")

                    if getdata == "ok" :
                        getdata = ""

                        for t in range( 0 , len(slct_menu)):
                            selctbox = g.driver0.execute_script("return document.getElementById('SLCTBOX').options[" + str(t) + "].selected")

                            if selctbox == True :
                                if return_kind == "str" :
                                    if getdata == "" :
                                        getdata = slct_menu[t]
                                    else:
                                        getdata = getdata + chr(9) + slct_menu[t]
                                else:
                                    if getdata == "" :
                                        getdata = str(t)
                                    else:
                                        getdata = getdata + chr(9) + str(t)
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return getdata

                    if i == 10 * TimeOut :
                        win=getid("all_win")
                        for i in range(0,len(g.all_win)):
                            if pos("選択してください。",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                                id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                                break
                        ctrlwin(id,"fs_min")
                        return -2 ##'///タイムアウト

                    sleep (100)
                    i = i + 1
            except:
                return ""


def Py_msg(msg,btn_type="btn_ok",sizex=500,sizey=100):
    '''
    msg             :メッセージ
    btn_type        :   "btn_ok"
                        "btn_ok_or_btn_no"
    '''
    import PySimpleGUI as sg

    menu_def = None
    ldict = {}

    sg.theme('Default1')
    menu_def = "g.menu_def='"+msg+"'"
    exec(menu_def, globals(), ldict)


    ## ウィンドウのレイアウトを定義
    if btn_type=="btn_ok":
        layout = [
        [sg.Text(g.menu_def,font=('メイリオ',12),size=(50, 1))],
        [sg.Button('Ok')]
        ]
    elif btn_type=="btn_ok_or_btn_no":
        layout = [
        [sg.Text(g.menu_def,font=('メイリオ',12),size=(50, 1))],
        [sg.Button('Ok'), sg.Button('NO')]
        ]

    ## Windowインスタンスの生成
    window = sg.Window("message", layout,size=(sizex,sizey))

    ## イベントループでWindow表示
    while True:
        event, values = window.read()

        ## Windowの「×」ボタンの処理と「終了」ボタンの処理
        if event == sg.WINDOW_CLOSED:
            Py_msg="close"
            break
        elif event == 'NO':
            Py_msg="btn_no"
            break
        elif event == 'Ok':
            Py_msg="btn_ok"
            break



    ## Windowを閉じる
    window.close()

    return Py_msg


def py_input_top():
    ##Py_msgを最前面にするためのスレッド用def
    sleep(100)

    while True:
        hwnd=getid("message","TkTopLevel")
        if hwnd > 0 :
            break
        sleep(100)

    ctrlwin(hwnd,"fs_active")

    keys = list()
    keys.append(g.vk_ctrl)
    keys.append(g.vk_a)
    sckey(hwnd,keys)


def Py_input(msg,default="",kind=1):
    '''
    msg             :メッセージ
    kind            :1      Ok / No
                    :2      Ok
    default         :デフォルトの入力値
    戻り値           :値
                    :False
    '''
    import PySimpleGUI as sg

    menu_def = None
    ldict = {}
    ldict2 = {}

    sg.theme('Default1')


    menu_def = "g.menu_def='"+msg+"'"
    exec(menu_def, globals(), ldict)

    menu_def2 = "g.menu_def2='"+default+"'"
    exec(menu_def2, globals(), ldict2)



    ## ウィンドウのレイアウトを定義
    if kind==1:
        layout = [
        [sg.Text(g.menu_def,font=('メイリオ',12),size=(50, 1))],
        [sg.Input(g.menu_def2,key='-INPUT-',font=('メイリオ',20))],
        [sg.Button('Ok'),sg.Button('No')]
        ]
    elif kind==2:
        layout = [
        [sg.Text(g.menu_def,font=('メイリオ',12),size=(50, 1))],
        [sg.Input(g.menu_def2,key='-INPUT-',font=('メイリオ',20))],
        [sg.Button('Ok')]
        ]


    th1 =Thread(py_input_top)

    ## Windowインスタンスの生成
    window = sg.Window("message", layout,size=(700,200),location=(10, 50),return_keyboard_events=True)

    ## イベントループでWindow表示
    while True:
        event, values = window.read()
        ##print(f"event='{event}', code={ord(event[0])}, values={values}\n")
        ## Windowの「×」ボタンの処理と「終了」ボタンの処理

        if event == sg.WINDOW_CLOSED:
            Py_msg=False
            break
        elif event == 'Ok':
            Py_msg=values['-INPUT-']
            break
        elif event == '\r':
            if values['-INPUT-'] == default :
                Py_msg=values['-INPUT-']
                break
            else:
                default=values['-INPUT-']

        elif event == 'No':
            Py_msg=False
            break

    ## Windowを閉じる
    window.close()

    Thread_end(th1)
    return Py_msg


def popupmenu(selectpopup,msg= "選択してください"):
    '''
    popupmenu (selectpopup)

    selectpopup             :選択するための配列

    戻り値；
        選ばれた配列位置
        g.popup_result に選ばれた文字が入る
    '''



    intext = "<html>\n"
    intext=intext+"<head>\n"
    intext=intext+"<title>PopupMenu</title>\n"
    intext=intext+"<script language='JavaScript'>\n"
    intext=intext+"function eventA(num){document.getElementById('event').innerText=num}\n"
    intext=intext+"</script>\n"
    intext=intext+"</head>\n"
    intext=intext+"<body width='30' height='60'>\n"

    intext=intext+msg+"<br><br>\n<select id='sel' onChange = 'eventA(this.value)'>\n<option value='-1'> </option>\n"
    for i in range(0,len(selectpopup)):
        intext=intext+"<option value='"+str(i)+"'>"+selectpopup[i]+"</option>\n"

    intext=intext+"</select><br>\n"
    intext=intext+"<span id='event' style='display:none'></span><br><br>\n"
    intext=intext+"<!--span id='event'></span>-->\n"

    intext=intext+"<input type='button' id='end' value='Cancel' onclick='eventA(this.id)'>\n"
    intext=intext+"</body>\n"
    intext=intext+"</html>\n"

    f1={}
    f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
    fput(f1, intext)
    fclose(f1)

    #Web
    try:
        g.driver0.get(g.CUR_DIR+"\Temp.html")
    except:
        g.driver0 = webdriver.Edge(service=g.service, options=options)
        g.driver0.get(g.CUR_DIR+"\Temp.html")

    g.driver0.set_window_size(500,400)
    win=getid("all_win")
    for i in range(0,len(g.all_win)):
        if pos("PopupMenu",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
            break
    ctrlwin(id,"fs_show")

    BusyWait(g.driver0)
    g.driver0.execute_script("return document.getElementById('sel').focus()")



    try:
        while True:
            getdata = g.driver0.execute_script("return document.getElementById('event').innerText")

            if getdata != "" and getdata != -1:
                if getdata == "end":
                    g.popup_result = -1
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("PopupMenu",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return -1
                else:
                    getnum = int(getdata)
                    g.popup_result = selectpopup[getnum]
                    ##g.driver0.quit()
                    win=getid("all_win")
                    for i in range(0,len(g.all_win)):
                        if pos("PopupMenu",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                            id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                            break
                    ctrlwin(id,"fs_min")
                    return getnum

            sleep(10)

    except:

        g.popup_result = -1
        return -1


def popupmenu2(table):
    '''
    table               :メニューの配列

    戻り値               :インデックス番号
    '''
    import PySimpleGUI as sg

    menu_def = None
    ldict = {}


    menu_def = "g.menu_def=[['Menu',["
    for i in range(0,len(table)):
        menu_def=menu_def+"'"+table[i]+"',"
    menu_def=menu_def+"]]]"

    ##print(menu_def)
    exec(menu_def, globals(), ldict)

    layout = [
        [sg.Menu(g.menu_def)]
    ]

    window = sg.Window('Menuの設定',layout,size=(150,50))

    getflag=False

    th1=Thread(pop_click)

    while True:
        event, value = window.read() #イベント入力を待つ
        #id=getid("Menuの設定","TkTopLevel")
        #clkitem(id,"TkChild","clk_btn",1)

        sleep(100)
        for i in range(len(table)):
            if event == table[i]:
                getdata=i
                getflag=True
                break

        if event is None:
            getdata=-1
            break

        if getflag == True:
            break

    window.close()

    Thread_end(th1)
    return getdata


def pop_click():
    ##popupmenu2を最前面にするスレッド用def
    while True:
        id=getid("Menuの設定","TkTopLevel")
        if id > 0:break
        sleep(10)

    ctrlwin(id,"fs_active")
    mmv(g.g_screen_w/2-50,g.g_screen_h/2+15)
    btn("left")


def popupmenu_pro(table):
    '''
    動かない
    '''
    import win32gui
    import win32con
    import ctypes

    user32 = ctypes.WinDLL("user32")

    SC_RESTORE = 0xF120
    MF_BYCOMMAND:int = 0x0
    TPM_RETURNCMD:int= 0x0100

    Handle = win32gui.CreatePopupMenu()
    for i in range(len(table)):
        getkey = table[i]
        Ret=user32.InsertMenuA(Handle, SC_RESTORE, MF_BYCOMMAND, (i + 1),getkey)


    hWnd = getid("Menuの設定","TkTopLevel")
    user32.SetForegroundWindow(hWnd) #最前面
    pos = win32gui.GetCursorPos()
    getdata=user32.TrackPopupMenu(Handle,0, pos[0], pos[1], 0,hWnd, 0)-1
    return getdata


def BusyWait(driver0):
    '''
    chrome/Edgeのビジー待ち
    driver0             :Seleniumのchrome/Edgeのドライバー
    '''
    sleep(100)
    while True:
        getdata = driver0.execute_script("return document.readyState")
        if getdata=="complete":break
        sleep(100)
    sleep(100)


def calender_th():
    ##カレンダーのスレッド用DEF

    sleep(100)
    while True:
        hwnd = getid("カレンダー","TkTopLevel")

        if hwnd > 0 :
            break
        sleep(10)

    sleep(10)

    rect = status(hwnd)
    clkitem(hwnd,"Button","clk_btn",2)

    mmv(rect[0]+230,rect[1]+48)
    btn("left")


def calender(kind=2):
    '''
    カレンダーを出す
    kind            :0      2023-04-01
                    :1      20230401
                    :2      2023/04/01
    '''

    import PySimpleGUI as sg

    th1 = Thread(calender_th)

    date = sg.Text(key='-text_date-', size=(20, 1))
    layout = [
                [date,sg.CalendarButton('日付選択',
                                          locale='ja_JP',
                                          format='%Y/%m/%d',
                                          month_names=["{:>2d}月".format(m) for m in range(1, 13)],
                                          key='-button_calendar-',
                                          font=('メイリオ',10),
                                          target='-text_date-')],
                ##[sg.Exit()]]
                [sg.Submit(button_text='OK',font=('メイリオ',10),pad=(0,10),size=(4,1),key='button_ok')]
              ]



    window = sg.Window('カレンダー', layout,return_keyboard_events=True)
    while True:
        event, values = window.read()

        if event == 'button_ok' or event == '\r':
            break
        elif event is sg.WIN_CLOSED:
            return False

    window.close()

    getdata =date.get()
    if kind == 0:
        getdata = replace(getdata,"/","-")
    elif kind == 1:
        getdata = replace(getdata,"/","")

    return getdata

    Thread_end(th1)


def calender_pro(msg="",kind=1):
    '''
    'kind           :0      2023-04-01
    'kind           :1      20230401
    'kind           :2      2023/04/01
    '''

    today = gettime(0)
    today = g.G_TIME_YY4+"-"+g.G_TIME_MM2+"-"+g.G_TIME_DD2

    web_str = "<html>\n"
    web_str = web_str+"<head>\n"
    web_str = web_str+"<title>カレンダー</title>\n"
    web_str = web_str+"<script language='JavaScript'>\n"
    web_str = web_str+"function eventA(getdata){\n"
    web_str = web_str+"   document.getElementById('eventA').innerText=getdata\n"
    web_str = web_str+"}\n"
    web_str = web_str+"</script>\n"
    web_str = web_str+"</head>\n"
    web_str = web_str+"<body width='30' height='60'>\n"
    web_str = web_str+msg+"<br>\n"
    web_str = web_str+"<input type='date' id='date1' name='calender' onchange='eventA(this.id)' value='" + today +"' /></td>\n"
    web_str = web_str+"<br><br><input id='ok' type='button' value = 'OK' onClick='eventA(this.id)'>\n"
    web_str = web_str+"  <input id='cancel' type='button' value = 'Cancel' onClick='eventA(this.id)'>\n"
    web_str = web_str+"<span id='eventA'  style='display:none;'></span></body>\n"
    web_str = web_str+"</html>"

    f1={}
    f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
    fput(f1, web_str)
    fclose(f1)

    ###########IE_msg###########
    try:
        g.driver0.set_window_size(g.g_screen_w * 2 / 5, 200)
    except:
        g.driver0 = webdriver.Edge(service=g.service, options=options)
        g.driver0.set_window_size(g.g_screen_w * 2 / 5, 200)
    ###########IE_msg###########


    g.driver0.get(g.CUR_DIR+"//Temp.html")
    sleep(200)
    BusyWait(g.driver0)
    g.driver0.refresh
    sleep(200)

    g.driver0.execute_script("document.getElementById('date1').focus()")

    try:
        while True:

            getdata = g.driver0.execute_script("return document.getElementById('eventA').innerText")

            if getdata == "date1" or getdata == "ok":
                getdata2  = g.driver0.execute_script("return document.getElementById('date1').value")

                ##g.driver0.quit
                win=getid("all_win")
                for i in range(0,len(g.all_win)):
                    if pos("カレンダー",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                        id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                        break
                ctrlwin(id,"fs_min")

                if kind == 0:
                    return getdata2
                elif kind == 1:
                    return replace(getdata2,"-","")
                elif kind == 2:
                    return replace(getdata2,"-","/")

            elif getdata == "cancel":
                ##g.driver0.quit
                win=getid("all_win")
                for i in range(0,len(g.all_win)):
                    if pos("カレンダー",g.all_win[i]) > 0 and pos("Chrome",g.all_win[i]) > 0:
                        id=copy(g.all_win[i],1,pos(",",g.all_win[i])-1)
                        break
                ctrlwin(id,"fs_min")
                return "cancel"

    except:
        return "cancel"


def exec_apli(app):
    '''
    app             :起動するアプリのパス
    j.exec_apli(r"\\tr\dfs\ELデバイス部\07_室課\03_プロ開室\神川\情報機器持出返却管理表\Temp.html")
    '''
    import os

    os.startfile(app, operation='open')


def doscmd(strA):
    '''
    strA            :dosコマンド
    '''
    import subprocess

    subprocess.run(strA,shell=True)


def set_printer(printername):
    '''
    printername             :"Microsoft Print to PDF"
                            :"\\hnps-sv04\TRJ-00037"
    j.doscmd("rundll32.exe printui.dll,PrintUIEntry /y /n "+chr(34)+"Microsoft Print to PDF"+chr(34))
    j.doscmd("rundll32.exe printui.dll,PrintUIEntry /y /n "+chr(34)+r"\\hnps-sv04\TRJ-00037"+chr(34))
    '''
    doscmd("rundll32.exe printui.dll,PrintUIEntry /y /n "+chr(34)+printername+chr(34))


def open_folder_dialog(dir_path=""):
    '''
    dir_path            :開くフォルダパス
    '''
    from tkinter import filedialog

    dir = dir_path
    fld = filedialog.askdirectory(initialdir = dir)
    return fld


def open_file_dialog(dir_path="",hyoudai=""):
    '''
    dir_path        :開くフォルダパス
    複数の時は"_"で結合

    '''

    from tkinter import filedialog

    typ = [('', '*')]
    dir = dir_path
    fle = filedialog.askopenfilenames(filetypes = typ, initialdir = dir , title = hyoudai)

    file_name=""
    cyc=0
    for f in fle:
        if cyc==0:
            file_name=f
        else:
            file_name=file_name+"_"+f

        cyc+=1
    return file_name


def formatA(num,keta):
    '''
    #ゼロ埋めする
    num             :元の数字
    keta            :ゼロ埋めして整える桁数
    '''
    getdata=format(num,"0"+str(keta))
    return getdata


def round(num,keta="0"):
    '''
    #小数点の桁合わせ
    num         :str型の数字
    keta        :"0"        整数にする
                :"0.1"      小数点1桁
                :"0.01"     小数点2桁
    '''
    from decimal import Decimal, ROUND_HALF_UP
    num= Decimal(str(num)).quantize(Decimal(keta), rounding=ROUND_HALF_UP)
    return num


def btn(kind,wheel=1):
    '''
    kind            :"left"
                    :"right"
                    :"middle"
                    :"wheel"   ホイール
    wheel           :ホイール操作　1:上　　-1:した
    '''
    import mouse

    #マウスの左ボタンをクリック
    if kind=="left":
        mouse.click('left')

    #マウスの右ボタンをクリック
    elif kind=="right":
        mouse.click('right')

    #マウスの中央ボタンをクリック
    elif kind=="middle":
        mouse.click('middle')

   #マウスのwheel
    elif kind=="wheel":
        mouse.wheel(wheel)


def mouse_position(hwnd=0):
    '''
    hwnd                :0  　絶対座標
                        :0以外　　hwndに対する相対座標
    マウスのポジションを得る
    ポジションは
    x位置　g.mouse_pos[0]
    y位置  g.mouse_pos[1]
    '''
    import mouse
    g.mouse_pos=Hairetu(2)
    g.mouse_pos=mouse.get_position()

    if hwnd == 0:
        g.mouse_pos=mouse.get_position()
        ###print(g.mouse_pos[0],g.mouse_pos[1])
        return g.mouse_pos
    else:
        rect=status(hwnd)
        posx=g.mouse_pos[0]-rect[0]
        posy=g.mouse_pos[1]-rect[1]
        g.mouse_pos=(posx,posy)
        print(posx,posy)

        return g.mouse_pos


def get_mouseposition():

    from ctypes import windll, Structure, c_long, byref

    class POINT(Structure):
        _fields_ = [
                        ("x", c_long),
                        ("y", c_long)
                    ]

    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return [pt.x,pt.y]


def mouse_cur():
    '''
    マウスの状態を得る
    戻り値             :マウスカーソルの種類（ハンドル）
    '''
    import win32gui
    flags, hcursor, pos = win32gui.GetCursorInfo()
    return hcursor


def mmv(posx,posy,kind="",posx2=0,posy2=0):
    '''
    posx        :移動するX位置
    posy        :移動するＹ位置
    kind        :""        移動
                :"drag"    ドラッグ
    posx2       :ドラッグするX位置
    posy2       :ドラッグするＹ位置
    '''

    import mouse

    if kind=="":
        mouse.move(posx,posy, absolute=True, duration=0.1)
    elif kind=="drag":
        mouse.drag(posx,posy,posx2,posy2,absolute = True,duration = 0.1)


def POFF(kind):
    '''
    kind        :"shutdown"
                :"logoff"
    '''
    import os

    if kind=="shutdown":
        os.system("shutdown /s /t 1")
    elif kind=="logoff":
        os.system("shutdown /l")


def mailsend(kenmei,Soushin,Bunsyo,kind=0,SoushinCC="",tenpu_files=0):#件名,送信先,文書,
    '''
    kenmei          :件名
    Soushin         :送信先
    Bunsyo          :メール本文
    kind            :0  表示させる
                    :1  メールを送信する
    SoushinCC       :CCの送信先
    tenpu_files     :添付ファイルの配列
    '''
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)

    mail.to = Soushin
    mail.cc = SoushinCC
    mail.bcc = ''
    mail.subject = kenmei
    mail.bodyFormat = 1
    mail.body = Bunsyo

    if tenpu_files != 0 :
        for i in range(0,len(tenpu_files)):
            mail.Attachments.Add(tenpu_files[i])

    if kind==0:
        mail.display(True)
    elif kind==1:
        mail.Send()


def mailseach(Folder1,Folder2,kind,word,read=0,kidoku=False):
    '''
    getdata=j.mailseach("teruyuki","受信トレイ","差出人","workflow",0,False)

    Folder1         :探すアカウント
    Folder2         :探すフォルダ
    kind            :"件名"
                    :"差出人"
                    :"本文"
    word            :"探す文字"
    read            :0 未読と既読
                    :1 未読
                    :2 既読
    kidoku          :False   既読にしない
                    :True    既読にする
    戻り値           :メールの情報（配列）
    '''
    import win32com.client
    from pathlib import Path

    stop_flag=0


    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    accounts = outlook.Folders
    move_file = list()

    ##print("root (アカウント数=%d)" % accounts.Count)
    for account in accounts:
        print("└ ",account)

        if pos(Folder1,str(account)) ==0:
            continue

        folders = account.Folders
        for folder in folders:
            print("  └ ",folder)

            if pos(Folder2,str(folder)) ==0:
                continue

            stop_flag=1
            find_cnt=0
            mailkoumoku=list()

            mails = folder.Items

            for mail in mails:
                find_flag=0

                try:
                    if read == 1 :
                        if mail.Unread != True:
                            continue
                    elif read == 2 :
                        if mail.Unread == True:
                            continue

                    if kind == "件名":
                        if pos(word,mail.subject)>0:
                            find_flag=1
                    elif kind == "差出人":
                        if pos(word,mail.sendername)>0:
                            find_flag=1
                        if pos(word,mail.senderEmailAddress)>0:
                            find_flag=1
                    elif kind == "本文":
                        if pos(word,mail.body)>0:
                            find_flag=1

                    if find_flag==1:
                        ##print(mail.subject)

                        attachment_file=""
                        for attachment in mail.Attachments:
                            if attachment_file == "" :
                                attachment_file = Path(attachment.FileName)
                            else:
                                attachment_file = attachment_file + "/" + Path(attachment.FileName)

                        mailkoumoku.append([mail.subject,mail.sendername,mail.receivedtime,mail.Unread,mail.body,attachment_file])
                        if kidoku==True:
                            mail.Unread = False
                            mail.save()
                        else:
                            mail.Unread = True
                            mail.save()
                        find_cnt+=1



                except:
                    pass

            if stop_flag==1:
                break
        if stop_flag==1:
            break

    return mailkoumoku


def mail_tenpfile_download(Folder1,Folder2,Save_path,kind,word,read=0,kidoku=False):
    '''
    getdata=j.mailseach("teruyuki","受信トレイ","差出人","workflow",0,False)

    Folder1         :探すアカウント
    Folder2         :探すフォルダ
    Save_path       :保存先
    kind            :"件名"
                    :"差出人"
                    :"本文"
    word            :"探す文字"
    read            :0 未読と既読
                    :1 未読
                    :2 既読
    kidoku          :False   既読にしない
                    :True    既読にする
    戻り値           :メールの情報（配列）
    '''
    import win32com.client
    from pathlib import Path

    stop_flag=0


    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    accounts = outlook.Folders
    move_file = list()

    ##print("root (アカウント数=%d)" % accounts.Count)
    for account in accounts:
        print("└ ",account)

        if pos(Folder1,str(account)) ==0:
            continue

        folders = account.Folders
        for folder in folders:
            print("  └ ",folder)

            if pos(Folder2,str(folder)) ==0:
                continue

            stop_flag=1
            find_cnt=0
            mailkoumoku=list()

            mails = folder.Items

            for mail in mails:
                find_flag=0

                try:
                    if read == 1 :
                        if mail.Unread != True:
                            continue
                    elif read == 2 :
                        if mail.Unread == True:
                            continue

                    if kind == "件名":
                        if pos(word,mail.subject)>0:
                            find_flag=1
                    elif kind == "差出人":
                        if pos(word,mail.sendername)>0:
                            find_flag=1
                        if pos(word,mail.senderEmailAddress)>0:
                            find_flag=1
                    elif kind == "本文":
                        if pos(word,mail.body)>0:
                            find_flag=1

                    if find_flag==1:
                        ##print(mail.subject)

                        attachment_file=""
                        for attachment in mail.Attachments:
                            if attachment_file == "" :
                                attachment_file = Path(attachment.FileName)
                            else:
                                attachment_file = attachment_file + "/" + Path(attachment.FileName)

                            attachment.SaveAsFile(path_join(Save_path,attachment_file))


                        mailkoumoku.append([mail.subject,mail.sendername,mail.receivedtime,mail.Unread,mail.body,attachment_file])
                        if kidoku==True:
                            mail.Unread = False
                            mail.save()
                        else:
                            mail.Unread = True
                            mail.save()
                        find_cnt+=1



                except:
                    pass

            if stop_flag==1:
                break
        if stop_flag==1:
            break

    return mailkoumoku


def mailmove(Folder1,Folder2,kind,word,Folder3,Folder4,kidoku=False):
    '''
    getdata=j.mailmove("teruyuki","受信トレイ","差出人","workflow","2023","受信",False)

    Folder1         :探すアカウント
    Folder2         :探すフォルダ
    kind            :"件名"
                    :"差出人"
                    :"本文"
    word            :"探す文字"
    Folder3         :移動するアカウント
    Folder4         :移動するフォルダ
    kidoku          :True    既読にする
                    :False　　既読にしない
    戻り値           :実施したメールの情報（配列）
    '''
    import win32com.client

    stop_flag=0


    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    accounts = outlook.Folders

    ##print("root (アカウント数=%d)" % accounts.Count)
    for account in accounts:
        print("└ ",account)

        if pos(Folder1,str(account)) ==0:
            continue

        folders = account.Folders
        for folder in folders:
            print("  └ ",folder)

            if pos(Folder2,str(folder)) ==0:
                continue

            stop_flag=1
            find_cnt=0
            mailkoumoku=list()
            mailmove=list()

            mails = folder.Items

            for mail in mails:
                find_flag=0

                try:
                    if kind == "件名":
                        if pos(word,mail.subject)>0:
                            find_flag=1
                    elif kind == "差出人":
                        if pos(word,mail.sendername)>0:
                            find_flag=1
                        if pos(word,mail.senderEmailAddress)>0:
                            find_flag=1
                    elif kind == "本文":
                        if pos(word,mail.body)>0:
                            find_flag=1

                    if find_flag==1:
                        print(mail.subject)
                        mailkoumoku.append([mail.subject,mail.sendername,mail.receivedtime,mail.Unread,mail.body])
                        mailmove.append(mail)

                        if kidoku==True:
                            mail.Unread = False
                            mail.save()
                        ##mail.move(accounts(Folder3).folders(Folder4))
                        find_cnt+=1
                except:
                    pass

            if stop_flag==1:
                break
        if stop_flag==1:
            break

    for i in range(0,len(mailmove)):
        mailmove[i].move(accounts(Folder3).folders(Folder4))
    return mailkoumoku


def up_dir(path):
    '''
    path            　:パス
    戻り値             :一つ上のディレクトリ
    '''
    import os
    return os.path.dirname(path)


def bunkatu_path(path):
    '''
    パスを分割する
    path1,path2 = bunkatu_path(path)

    path1       :ひとつ上の階層のパス
    path2       :ファイル名　や　最下位のフォルダ名
    '''
    import os

    path1=os.path.dirname(path)
    path2=os.path.basename(path)

    return path1,path2


def path_join(*path_str):
    '''
    パスを連結する
    path_str            :path1,path2,path3,…………

    '''
    import os

    path_table=list()
    path_table=path_str
    return os.path.join(*path_table)


def End():
    '''
    Pythonの終了
    '''
    import sys

    sys.exit()


def baloon(msg,title="通知"):
    '''
    バルーンチップメッセージを出す

    msg     :メッセージ
    title   :タイトル
    '''
    from plyer import notification
    notification.notify(title=msg,message=title,app_name="python",)


def ex_OPEN(path="",visible=True):
    '''
    ex,wb=ex_OPEN(path)

    戻り値：   excel:アプリケーションOBJ
               wb:ワークブックOBJ


    '''
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    File_name1,File_name=bunkatu_path(path)
    if path != "":
        excel.Workbooks.Open(path)
        wb=excel.Workbooks(File_name)
    else:#新規
        wb = excel.Workbooks.Add()

    return excel,wb


def ex_OPEN2(path="",visible=True):
    '''
    ##既存のEXCELから立ち上げる
    ex,wb=ex_OPEN2(path)

    戻り値：   excel:アプリケーションOBJ
               wb:ワークブックOBJ

    '''

    import win32com.client
    excel = win32com.client.GetObject(Class="Excel.Application")
    excel.Visible = visible

    File_name1,File_name=bunkatu_path(path)

    if path != "":
        Flag=False
        for iv in range(0,len(excel.Workbooks)):
            if excel.Workbooks[iv] == File_name :
                wb=excel.Workbooks(File_name)
                Flag=True
                break
        if Flag==False:
            wb=excel.Workbooks.Open(path)

    else:#新規
        wb = excel.Workbooks.Add()

    return excel,wb


def ex_GETCELL(ws,ex_kind,ex_pos=1):
    '''

'    ex_GETCELL(ws,ex_kind,[ex_pos])

'        ws；オブジェクト
'        ex_kind；取り出す種類
'            ex_endrow          ；最終行
'            ex_endrow_ad       ；最終行アドレス
'            ex_endcell         ；最終列
'            ex_endcell_ad      ；最終列アドレス

'            一塊の入力範囲に対して最終行などを取得する場合
'            ex_endrow2          ；最終行
'            ex_endrow_ad2       ；最終行アドレス
'            ex_endcell2         ；最終列
'            ex_endcell_ad2      ；最終列アドレス

'            ex_sel_row         ；選択エリアの行数
'            ex_sel_cell        ；選択エリアの列数
'            ex_sel_row_ad1     ；選択エリアの先頭行
'            ex_sel_cell_ad1    ；選択エリアの先頭列
'            ex_sel_row_ad2     ；選択エリアの最終行
'            ex_sel_cell_ad2    ；選択エリアの最終列
'        ex_pos；
'                探す行　または　列
'                一塊に対して行う場合はアドレスを指定：　"A5"など

'        戻り値；値


    '''

    import win32com.client

    xlUp=-4162
    xlToLeft=-4159
    xlDown=-4121
    xlToRight=-4161


    if ex_kind == "ex_endrow":
       return ws.Cells(ws.Rows.Count, ex_pos).End(xlUp).Row #''//###########最終行###########

    elif ex_kind == "ex_endrow_ad":
       getdata = ws.Cells(ws.Rows.Count, ex_pos).End(xlUp).Address #''//###########最終行###########
       return replace(getdata, "$", "")

    elif ex_kind == "ex_endcell":
       return ws.Cells(ex_pos, ws.Columns.Count).End(xlToLeft).Column #''//###########最終列###########

    elif ex_kind == "ex_endcell_ad":
        getdata = ws.Cells(ex_pos, ws.Columns.Count).End(xlToLeft).Address #''//###########最終列###########
        return replace(getdata, "$", "")

    elif ex_kind == "ex_endrow2":
       return ws.Range(ex_pos).End(xlDown).Row #''//###########最終行###########

    elif ex_kind == "ex_endrow_ad2":
       getdata = ws.Range(ex_pos).End(xlDown).Address #''//###########最終行###########
       return replace(getdata, "$", "")

    elif ex_kind == "ex_endcell2":
       return ws.Range(ex_pos).End(xlToRight).Column #''//###########最終列###########

    elif ex_kind == "ex_endcell_ad2":
        getdata = ws.Range(ex_pos).End(xlToRight).Address #''//###########最終列###########
        return replace(getdata, "$", "")

    elif ex_kind == "ex_sel_row":
        return ws.Selection.Rows.Count

    elif ex_kind == "ex_sel_cell":
        return ws.Selection.Columns.Count

    elif ex_kind == "ex_sel_row_ad1":
        return ws.Selection.Cells(1).Row

    elif ex_kind == "ex_sel_cell_ad1":
        return ws.Selection.Cells(1).Column

    elif ex_kind == "ex_sel_row_ad2":
        return ws.Selection.Cells(ws.Selection.Count).Row

    elif ex_kind == "ex_sel_cell_ad2":
        return ws.Selection.Cells(ws.Selection.Count).Column


def ex_close(excel,wb,ex_kind,kind2=""):
    '''
    excel       :applicationオブジェクト
    wb          :workbookオブジェクト
    ex_kind     :"Close"  保存しない
                :"Save"   上書き保存する
                :"保存先"  名前を付けて保存する
    kind2       :"Quit"   excelを終わる
    '''
    import win32com.client

    if ex_kind == "Close":
        wb.Close(False)

        if kind2=="Quit":
            excel.Quit()

    elif ex_kind == "Save":
        wb.Save()
        wb.Close(False)

        if kind2=="Quit":
            excel.Quit()

    else:
        wb.SaveAs(ex_kind)
        wb.Close(False)

        if kind2=="Quit":
            excel.Quit()



def ex_PDF(ws,save_path):
    '''
    ws                  :ワークシートオブジェクト
    save_path           :保存先
    '''
    ws.ExportAsFixedFormat(0, save_path)



def Thread(sub,h1=0):
    '''
    shread名=Thread(関数名,引数)
    引数は必ずh1にすること

    def AAA(h1):
        j.msg(h1)

    th1=j.Thread(AAA,"hohoho")
    '''
    import threading
    if h1 != 0:
        thread = threading.Thread(kwargs={'h1':h1},target=sub)
    else:
        thread = threading.Thread(target=sub)
    thread.start()
    return thread


def Thread_end(thread):
    import threading
    thread.join()


def CHKIMG(img,x=0,y=0,w="",h="",display=False):
    '''
    kazu=j.CHKIMG('test.png',0,0,500,500,False)
    j.mmv(g.CHK_X[0],g.CHK_Y[0])

    img:            :検出画像のパス
    x,y,w,h         :検索範囲
    戻り値           :画像の数
                    :g.CHK_X[]      X座標
                    :g.CHK_Y[]      y座標
    '''

    ##背景のキャプチャ
    if w=="":w=g.g_screen_w
    if h=="":h=g.g_screen_h
    play_area = (x, y, w, h) #x, y, +x, +y
    backimg = get_screen(play_area)

    ##ターゲット画像
    targetimg = cv2.imread(img)
    targetimg = cv2.cvtColor(targetimg,cv2.COLOR_BGR2RGB)
    ##show(targetimg)

    pos = template_match(backimg, targetimg, visible=display)
    ##print(pos)

    g.CHK_X=list()
    g.CHK_Y=list()

    for i in range(0,len(pos)):
        g.CHK_X.append(pos[i][0]+x)
        g.CHK_Y.append(pos[i][1]+y)

    return len(pos)


def peekcolor():
    '''
    R,G,B=peekcolor()
    '''

    import win32gui
    import win32api
    import ctypes

    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()

    desktop = win32gui.GetDesktopWindow()
    document = win32gui.GetDC(desktop)

    ## カーソル位置の取得
    cursor_pos = win32api.GetCursorPos()

    ## 指定したPixelのRGB値の取得
    pixel_color = win32gui.GetPixel(
        document, cursor_pos[0], cursor_pos[1]
    )

    ## 10進数表記のRGB値に変換
    rgb = [pixel_color >> 8*i & 0xff for i in range(3)]

    ##print("RGB({0},{1},{2})".format(*rgb))

    return rgb


def color_choose(kind=0):
    '''
    R,G,B=j.color_choose(0)         R,G,Bで返る
    R=j.color_choose(1)             #16進数で返る
    '''

    import tkinter
    from tkinter import colorchooser

    RGB,RGB16=colorchooser.askcolor() #colorchooser呼び出し

    if kind==0:
        return RGB[0],RGB[1],RGB[2]
    else:
        return RGB16


def OCR(x=0,y=0,w="",h="",val=150,kind=0,kakudai=1):
    '''
    table=j.OCR(0,0,400,400)
    x,y,w,h         :画像認識範囲
    val             :濃さでの閾値
    kind            :0      通常のOCR
                    :1      座標を取得する
    kakudai         :画像の拡大率
    '''

    import pyautogui

    if w=="":w=g.g_screen_w
    if h=="":h=g.g_screen_h

    img2 = pyautogui.screenshot('screenshot.png', region=(x,y,w,h))
    data_list = render_doc_text('screenshot.png',val,kind,kakudai)
    return data_list


def drop_path(msg_str="",sizex=400,sizey=50):
    '''
    msg_str               :表示する文字　　ドロップしてください
    戻り値                 :×close         False
                          :ドロップファイルのパス
    '''

    def drop(event):
        text.set(event.data)
        g.drop_path=event.data
        root.quit()
        root.destroy()

    def quit_me(root_window):
        g.drop_path=False
        root_window.quit()
        root_window.destroy()


    # メインウィンドウの生成
    root = TkinterDnD.Tk()
    root.geometry(str(sizex)+"x"+str(sizey))
    root.title("ドラッグ&ドロップ")

    # StringVarのインスタンスを格納するウィジェット変数text
    text = tk.StringVar(root)
    if msg_str == "":
        text.set("ファイルをドロップしてください")
    else:
        text.set(msg_str)

    # Lavelウィジェットの生成
    label = ttk.Label(root,textvariable=text,font=("MSゴシック", "12", "bold"))
    label.drop_target_register(DND_FILES)
    label.dnd_bind("<<Drop>>", drop)

    # ウィジェットの配置
    label.grid(row=0, column=0, padx=10)
    root.protocol("WM_DELETE_WINDOW",lambda :quit_me(root))
    root.mainloop()

    if g.drop_path != False:
        if pos("{",g.drop_path) > 0 :
            g.drop_path=replace(g.drop_path,"{","")
            g.drop_path=replace(g.drop_path,"}","")
    return g.drop_path


def Open_Gazou(file):

    import cv2
    img = cv2.imread(file)
    cv2.imshow(file, img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    return ""


def DenshiIn_create(moji_Upper="",moji_Lower="",days=0):
    '''
    path=DenshiIn_create("主幹","神川"):
    path            :作成したスタンプのパス
    '''

    from PIL import Image, ImageDraw, ImageFont
    import os
    import time


    base_W, base_H = (320,320)
    ##moji_Upper = "部署X"
    ##moji_Lower = "氏 名"
    yohaku = 10

    image = Image.new('RGBA', (base_W, base_H), 'white')
    image.putalpha(0)
    d = ImageDraw.Draw(image)

    d.ellipse([(yohaku, yohaku), (base_W - yohaku, base_H - yohaku)], outline='red', width=6)  # 円を描画

    #はんこ上部の文字を入れる
    adjust = 9
    hight = (( (base_H-yohaku*2) / 3 ) * 1) + yohaku
    d.line([(0+yohaku+adjust,hight),(base_W-yohaku-adjust,hight)],fill='red',width=6) # よこ線を入れる

    fontpath = "C:\Windows\Fonts\msgothic.ttc"

    if len(moji_Upper) < 4:
        fontsize = 60
    elif len(moji_Upper) ==4 :
        fontsize = 48
    elif len(moji_Upper) <=6 :
        fontsize = 36
    elif len(moji_Upper) >= 7 :
        fontsize = 24

    font = ImageFont.truetype(fontpath, fontsize)
    moji_W=fontsize*len(moji_Upper)
    moji_H = fontsize



    moji_W = (base_W/2) - (moji_W/2)
    moji_H = hight - moji_H - yohaku
    d.text(( moji_W,moji_H), moji_Upper, font=font, fill='red')

    hight = (( (base_H-yohaku*2) / 3 ) * 2) + yohaku
    d.line([(0+yohaku+adjust,hight),(base_W-yohaku-adjust,hight)],fill='red',width=6) # よこ線を入れる

    #日付を入れる
##    now = time.localtime()
##    date = time.strftime('%y/%m/%d', now)

    if days ==0:
        gettime(days)
        date=g.G_TIME_YY2+"/"+g.G_TIME_MM2+"/"+g.G_TIME_DD2
    else:
        date=days

    fontpath = "C:\Windows\Fonts\msgothic.ttc"
    font = ImageFont.truetype(fontpath, 60)
    try:
        moji_W, moji_H = d.textsize(date, font=font)
    except:
        moji_W = 240
        moji_H=59

    moji_W = (base_W/2) - (moji_W/2)
    moji_H = (base_H/2) - (moji_H/2)
    d.text(( moji_W,moji_H), date, font=font, fill='red')


    #はんこ下部の文字を入れる
    fontpath = "C:\Windows\Fonts\msgothic.ttc"
    if len(moji_Lower) < 4:
        fontsize = 60
    elif len(moji_Lower) ==4 :
        fontsize = 48
    elif len(moji_Lower) <=6 :
        fontsize = 36
    elif len(moji_Lower) >= 7 :
        fontsize = 24

    font = ImageFont.truetype(fontpath, fontsize)
    moji_W=fontsize*len(moji_Lower)
    moji_H = fontsize



    moji_W = (base_W/2) - (moji_W/2)
    moji_H = hight + yohaku
    d.text(( moji_W,moji_H), moji_Lower, font=font, fill='red')

    image.save(path_join(g.CUR_DIR,"stamp.png"))
    return path_join(g.CUR_DIR,"stamp.png")


def DenshiIn_Paste(stamp_path,path1,path2,zahyou):
    '''
    ##一括で電子印を押す
    stamp_path              :スタンプのパス
    path1                   :PDFのパス
    path2                   :実施後のPDFのパス
    zahyou                  :スタンプの貼り付け場所        　zahyou[スタンプを押す個数][0] X
    zahyou                                              zahyou[スタンプを押す個数][1] Y
                                                        zahyou[スタンプを押す個数][2] サイズ
    戻り値                   :実施後のPDFのパス（配列）
    '''

    from pdfrw import PdfReader
    from pdfrw.buildxobj import pagexobj
    from pdfrw.toreportlab import makerl
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader

    import glob
    import os
    import PyPDF2
    from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer3.converter import PDFPageAggregator
    from pdfminer3.pdfpage import PDFPage
    from pdfminer3.layout import LAParams, LTTextContainer

    import pandas as pd

    from PIL import Image
    import io

    pdf_path=list()

    pdf_kazu = getdir(path1,"*.pdf")
    pdf_list = list()
    for i in range(0,pdf_kazu):
        pdf_list.append(g.getdir_files[i])
        print(pdf_list[i])

    for i in range(0,pdf_kazu):
        pdf_file = path_join(path1,pdf_list[i])
        out_file = path_join(path2,pdf_list[i])

        pdf_path.append(out_file)

        stamp_file = stamp_path

        img = ImageReader(stamp_file)
        ##j.Open_Gazou(stamp_file)

        reader = PdfReader(pdf_file)
        # ページ数を取得する
        Num_pages=len(reader.pages)
        print(len(reader.pages))

        pdf = PdfReader(pdf_file, decompress=False).pages
        cvs = canvas.Canvas(out_file)

        for ii in range(Num_pages):
            page = pagexobj(pdf[ii])
            cvs.doForm(makerl(cvs, page))

            for m in range(0,len(zahyou)):
                size=float(zahyou[m][2])
                size=size/2
                cvs.drawImage(img, (zahyou[m][0]-size)*mm, (zahyou[m][1]-size)*mm, zahyou[m][2]*mm, zahyou[m][2]*mm, preserveAspectRatio=True,mask='auto')          ##受領
            cvs.showPage()

        cvs.save()

    return pdf_path


def DenshiIn_Paste2(stamp_path,path1,path2,PDF_table,zahyou,do_page=0,Uwagaki=False):
    '''
    ##実施するPDFのみ電子印を押す
    stamp_path              :スタンプのパス
    path1                   :PDFのパス
    path2                   :実施後のPDFのパス
    PDF_table               :実施するPDFの配列（ファイル名のみの配列)
    zahyou                  :スタンプの貼り付け場所
                                zahyou[スタンプを押す個数][0] X
                                zahyou[スタンプを押す個数][1] Y
                                zahyou[スタンプを押す個数][2] サイズ
    do_page                 :0      すべてのページ
                            :0以外　実施するページ
    Uwagaki                 :True   上書きする   False   別名保存
    戻り値                   :実施後のPDFのパス（配列）
    '''



    from pdfrw import PdfReader
    from pdfrw.buildxobj import pagexobj
    from pdfrw.toreportlab import makerl
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader

    import glob
    import os
    import PyPDF2
    from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer3.converter import PDFPageAggregator
    from pdfminer3.pdfpage import PDFPage
    from pdfminer3.layout import LAParams, LTTextContainer

    import pandas as pd

    from PIL import Image
    import io

    pdf_path=list()


    pdf_kazu = len(PDF_table)
    pdf_list = list()
    for i in range(0,pdf_kazu):
        pdf_list.append(PDF_table[i])
        print(pdf_list[i])

    for i in range(0,pdf_kazu):
        pdf_file = path_join(path1,pdf_list[i])

        if Uwagaki == True :
            out_file = path_join(path2,pdf_list[i])
        else:
            out_file = path_join(path2,replace(pdf_list[i],".pdf","_印.pdf"))

        pdf_path.append(out_file)

        stamp_file = stamp_path

        img = ImageReader(stamp_file)
        ##j.Open_Gazou(stamp_file)

        reader = PdfReader(pdf_file)
        # ページ数を取得する
        Num_pages=len(reader.pages)
        print(len(reader.pages))

        pdf = PdfReader(pdf_file, decompress=False).pages
        cvs = canvas.Canvas(out_file)

        if do_page == 0 :
            for ii in range(Num_pages):
                page = pagexobj(pdf[ii])
                cvs.doForm(makerl(cvs, page))

                for m in range(0,len(zahyou)):
                    size=float(zahyou[m][2])
                    size=size/2
                    cvs.drawImage(img, (zahyou[m][0]-size)*mm, (zahyou[m][1]-size)*mm, zahyou[m][2]*mm, zahyou[m][2]*mm, preserveAspectRatio=True,mask='auto')          ##受領
                cvs.showPage()
        else:
            for ii in range(Num_pages):

                page = pagexobj(pdf[ii])
                cvs.doForm(makerl(cvs, page))

                if do_page == (ii+1) :
                    for m in range(0,len(zahyou)):
                        size=float(zahyou[m][2])
                        size=size/2
                        cvs.drawImage(img, (zahyou[m][0]-size)*mm, (zahyou[m][1]-size)*mm, zahyou[m][2]*mm, zahyou[m][2]*mm, preserveAspectRatio=True,mask='auto')          ##受領
                cvs.showPage()
        cvs.save()

    return pdf_path


def PDF_Bunkatu(input_file,kind=0,wariai=0.5):
    '''
    input_file              :分割するＰＤＦファイルのパス
    kind                    :0  縦分割   1  横分割
    wariai                  :分割比
    戻り値                   :分割されたファイルパス1,分割されたファイルパス2
    '''

    import PyPDF2
    import copy
    from decimal import Decimal

    path1,path2=bunkatu_path(input_file)
    output_file1 = path_join(path1,"1-"+path2)
    output_file2 = path_join(path1,"2-"+path2)

    pdf_reader = PyPDF2.PdfReader(input_file)
    pdf_writer1 = PyPDF2.PdfWriter()
    pdf_writer2 = PyPDF2.PdfWriter()


    for i in range(len(pdf_reader.pages)):
        ## 同じページのオブジェクトを２つ用意
        p1 = pdf_reader.pages[i]
        p2 = copy.copy(p1)
        ## 原稿の左下と右上の座標を取得（用紙サイズ）
        x0 = p1.mediabox.left
        y0 = p1.mediabox.bottom
        x1 = p1.mediabox.right
        y1 = p1.mediabox.top

        ## 左右に分割して切り抜く領域の座標を計算
        p1_lower_left = (x0, y0)
        p1_upper_right = (int((x0 + x1)*Decimal(1-wariai)), y1)
        p2_lower_left = (int((x0 + x1)*Decimal(1-wariai)), y0)
        p2_upper_right = (x1, y1)


        if kind==0:
            ## 縦長の場合は上下で分割するように変える
            p1_upper_right = (x1,int((y0 + y1)*Decimal(1-wariai)))
            p2_lower_left = (x0, int((y0 + y1)*Decimal(1-wariai)))

        ## 切り抜く領域（cropBox）の設定
        p1.cropbox.lower_left = p1_lower_left
        p1.cropbox.upper_right = p1_upper_right
        p2.cropbox.lower_left = p2_lower_left
        p2.cropbox.upper_right = p2_upper_right

        ##print(p1_lower_left,p1_upper_right,p2_lower_left,p2_upper_right)

        ## 縦長の場合は上,下の順に並び替える（不要な場合はこの２行は削除）
        if kind==0:
            p1, p2 = p2, p1
        ## 出力用のオブジェクトに２ページ分を追加
        pdf_writer1.add_page(p1)
        pdf_writer2.add_page(p2)

    ## ファイルに出力
    with open(output_file1, mode="wb") as f:
        pdf_writer1.write(f)
    with open(output_file2, mode="wb") as f:
        pdf_writer2.write(f)

    return output_file1,output_file2


def PDF_set_password(src_path, dst_path, user_password, owner_password=None):
        '''
        src_path            :PDFファイルパス
        dst_path            :新規PDFファイルパス
        user_password       :ユーザーパスワード
        owner_password      :オーナーパスワード
        '''

        import pypdf
        src_pdf = pypdf.PdfReader(src_path)
        dst_pdf = pypdf.PdfWriter()
        dst_pdf.clone_reader_document_root(src_pdf)

        d = {key: src_pdf.metadata[key] for key in src_pdf.metadata.keys()}
        dst_pdf.add_metadata(d)

        dst_pdf.encrypt(user_password, owner_password)
        dst_pdf.write(dst_path)



def Lockhard(hwnd,kind):
    '''
    ##対象ウィンドウに対しキーをロックする
    hwnd                :対象のウィンドウハンドル
    kind                :True       キー操作を有効にする
                        :False      キー操作を無効にする
    '''
    import win32gui
    import win32con
    import ctypes

    sleep(10)
    enabled=kind
    win32gui.EnableWindow(hwnd,enabled)
    sleep(10)


def Lockhard2(kind):
    '''
    ##キー入力をロック
    kind            :True   ロックする
                    :False  ロック解除
    '''

    import keyboard

    sleep(10)
    if kind == True:


        AllKeySet = {
        'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
        'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
        '!','"','#','$','%','&',"'",'(',')','=','~','|','-','^','\\','@','[',';',':',']',',','`','{','}','<','>','?','_',
        '0','1','2','3','4','5','6','7','8','9','.','+','*','/',"num lock",
        "esc","f1","f2","f3","f4","f5","f6","f7","f8","f9","f10","f11","f12","print screen","scroll lock","pause","sys req",
        "insert","delete","home","end","page up","page down",
        "right","down","left","up",
        "半角/全角","tab","caps lock","shift","ctrl","left windows","alt","無変換","space",
        "変換","ひらがな","right alt","right windows","menu","right ctrl","right shift","enter","backspace"
        }


        for key in AllKeySet:
            keyboard.block_key(key)

    else:
        keyboard.unhook_all()

    sleep(10)


def Lockhard3(kind):
    '''
    ##キーとマウスをロックする
    kind            :True   ロックする
                    :False  ロック解除
    '''

    import win32gui
    import win32con
    import ctypes

    sleep(10)
    user32 = ctypes.WinDLL("user32")
    user32.BlockInput(kind)
    sleep(10)














'''
Arduino側

void setup() {
  Serial.begin(115200);
}


String myString;

void loop() {
  while(true){
    if(Serial.available()>=1) break;
    delayMicroseconds(100);
  }
  myString = Serial.readStringUntil(0x0a);//改行まで
  Serial.println(myString);
}

'''

def Serial_port():
    import serial
    import serial.tools.list_ports

    for port in serial.tools.list_ports.comports():
        print(f'# {port = } / {port.device = }')

def Serial_begin(com,late=115200):
    '''
    Port = j.Serial_begin("com6",115200)
    com         :comポート  "com6"
    late        :115200など
    '''

    import serial
    Serial_Port=serial.Serial(port=com,baudrate=late,parity= 'N')
    return Serial_Port

def Serial_write(Port,data):
    data = str(data)+"\n"
    data=data.encode('utf-8')
    Port.write(data)

def Serial_read(Port):
    sleep(5)
    data=Port.read_all()
    data=data.strip()
    data=data.decode('utf-8')
    return data

def Serial_close(Port):
    Port.close()








if __name__ == '__main__':
    main()
