
import global_value as g

import getpass
import os
import sys
import tkinter as tk
import time
import pyautogui

#web
from selenium import webdriver
import chromedriver_autoinstaller
chromedriver_autoinstaller.install()

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ['enable-automation'])

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
g.keytable["vk_capital"]=0x14
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
g.keytable["k_numpad1"]=0x61
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
g.keytable["k_f11"]=0x7A
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
g.keytable["vk_sepa"]=0x6C
g.keytable["vk_sub"]=0x6D
g.keytable["vk_dec"]=0x6E
g.keytable["vk_div"]=0x6F
g.keytable["vk_num"]=0x90
g.keytable["vk_Scroll"]=0x91
g.keytable["vk_kake1"]=0xBA
g.keytable["vk_tasu1"]=0xBB
g.keytable["vk_kanma"]=0xBC
g.keytable["vk_hiku"]=0xBD
g.keytable["vk_ten"]=0xBE
g.keytable["vk_que"]=0xBF
g.keytable["vk_at"]=0xC0
g.keytable["vk_kakko1"]=0xDB
g.keytable["vk_en"]=0xDC
g.keytable["vk_kakko2"]=0xDD
g.keytable["vk_kara"]=0xDE
g.keytable["vk_sura"]=0xE2
g.keytable["vk_caps"]=0xF0
g.keytable["vk_kana"]=0xF2
g.keytable["k_han"]=0xF3





def main():
    print("")
##    pass






def Hairetu(num1,num2=1):
    '''
    num1      1次元要素数
    num2      2次元要素数
    return    配列

    '''
    table=[[0 for i in range(num2)] for ii in range(num1)]
    return table



def Hairetu_sort(hairetu,kind=""):
    '''
    hairetu     配列
    kind        "" 昇順
    　　　　　　"reverse"  降順
    '''
    if kind=="":
        hairetu2=sorted(hairetu)

    elif kind=="reverse":
        hairetu2=sorted(hairetu,reverse=True)

    return hairetu2


def getid(g_title,g_class=None):
    '''
    getid(タイトル、クラス）
    g_title:            タイトル名
    g_class:            クラス名（省略可）
    '''

    from win32gui import GetWindowText,GetClassName,GetForegroundWindow,GetCursorPos

    if g_title=="get_active_win":
        hwnd=GetForegroundWindow()
        g_title=GetWindowText(hwnd)
        g_class=GetClassName(hwnd)
        print(g_title+","+g_class)
        return hwnd
    elif g_title=="get_frompoint_win":
        myx,myy=win32gui.GetCursorPos()
        hwnd = win32gui.WindowFromPoint((myx, myy))
        GA_ROOT=2
        hwnd = win32gui.GetAncestor(hwnd, GA_ROOT)
        g_title=GetWindowText(hwnd)
        g_class=GetClassName(hwnd)
        print(g_title+","+g_class)
        return hwnd
    else:
        hwnd = win32gui.FindWindow(g_class,g_title)
        if hwnd != 0:
            g_title=GetWindowText(hwnd)
            g_class=GetClassName(hwnd)
            print(g_title+","+g_class)
        return hwnd


def ACW(hwnd,x,y,wx=0,wy=0):
    '''
    hWnd                   ウィンドウハンドル
    x                      頂点座標X
    y                      頂点座標Y
    wx                     幅
    wy                     高さ
    '''

    import win32gui

    if wx==0 or wy==0:
        get_gamen_size()
        wx=g.g_screen_w
        wy=g.g_screen_h

    win32gui.MoveWindow(hwnd, x, y, wx, wy,1)




def ctrlwin(hwnd,meirei):
    #hWnd                   ウィンドウハンドル
    #meirei                 "fs_topmost"
    #                       "fs_notopmost"
    #                       "fs_active"
    #                       "fs_show"
    #                       "fs_min"
    #                       "fs_close"

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



#親ウィンドウから子ウィンドウを取得する
def get_sub_hWnd(p_handle):
    import win32gui
    import win32con
    import array
    import ctypes
    import os.path

    # 親ウィンドウハンドル(識別番号)とそのクラス名、ハンドル名を取得
    #p_handle = win32gui.FindWindow(None,"XXX") # アプリケーションのタイトル文字列など入力
    p_class_name = win32gui.GetClassName(p_handle)
    p_handle_name = win32gui.GetWindowText(p_handle)
    handles_dict = {str(p_handle): [p_class_name,p_handle_name]}
    #print('p',hex(p_handle),p_class_name,p_handle_name)

    # 親ウィンドウハンドルの子ウィンドウハンドルを配列に格納
    c_handles = array.array("i")  # 子ウィンドウハンドルは複数あることを想定し配列を作成
    win32gui.EnumChildWindows(p_handle,lambda handle,list: list.append(handle),c_handles)

    # 子ウィンドウハンドルとそのクラス名、ハンドル名を列挙
    g.hWnd_list={}
    str_caption=""
    user32 = ctypes.WinDLL("user32")
    get_str=ctypes.create_unicode_buffer(256)

    for hWnd in c_handles:
        c_class_name = win32gui.GetClassName(hWnd)
        c_handle_name = win32gui.GetWindowText(hWnd)

        length = user32.GetWindowTextLengthW(hWnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hWnd, buff, length + 1)
        strLen = win32gui.SendMessage(hWnd, 0xD, 256, get_str) #'''''読み取り
        #print(buff.value,get_str.value)

        print('c',hWnd,c_class_name,c_handle_name,get_str.value)
        handles_dict[hWnd] = [c_class_name,c_handle_name]

        if hash_data_out(g.hWnd_list,"hash_exists",c_class_name)==False:
            g.hWnd_list[c_class_name]=str(hWnd)
        else:
            getdata=g.hWnd_list[c_class_name]
            g.hWnd_list[c_class_name]=getdata+"_"+str(hWnd)

    for k in g.hWnd_list.keys():
        print(k,g.hWnd_list[k])



def getstr(hwnd,kind="EDIT",num=1):
    #hwnd           0:クリップボード
    #               0以外:親ハンドル
    #kind           クラス名
    #num            同じクラス名の何個目か


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
    #hwnd           0:クリップボード
    #               0以外:親ハンドル
    #Sendmoji       送る文字
    #kind           クラス名
    #num            同じクラス名の何個目か


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
    #hwnd           親ウィンドウハンドル
    #k_class        子ウィンドウのクラス名
    #do_kind        "clk_btn","clk_combo_select","clk_list_select","clk_check"
    #num            同じクラス名の何個目か
    #do_name        select時の選択するアイテム名

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
    elif do_kind == "clk_combo_select":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0x14D, 0, do_name) #''''選択
    elif do_kind == "clk_list_select":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0x186, 0, do_name) #''''クリック
    elif do_kind == "clk_check":
        user32.SendMessageA(get_hwnd_num, 0x6, 0, 0) #''''ｱｸﾃｨﾌﾞ
        user32.SendMessageA(get_hwnd_num, 0xF0, 0, 0) #''''クリック





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
                key=key+"_"+k


    return key



def sckey(id,key):
    '''
    id:ハンドル
    key:送るキーを列挙（最初だけ押しっぱなしで実施する）

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







def sleep(ms):
    #sleep("ミリ秒")

    time.sleep(ms/1000)


def getuser():
    getdata = getpass.getuser()
    return getdata



def get_computer_name():
    import socket

    # コンピュータ名を取得
    host = socket.gethostname()
    print(host)
    return host



def getdir(path,name):
    #path            探すパス
    #name            ”/”フォルダ検索
    #                ファイルは"*.py"などで指定

    import glob
    import os

    g.getdir_files={}
    folderfile = os.listdir(path)

    #フォルダの場合
    if name=="/":
        folder = [f for f in folderfile if os.path.isdir(os.path.join(path, f))]
        p=0
        for name in folder:
            #print(name)
            g.getdir_files[p]=name
            p=p+1

        return p

    #ファイルの場合
    else:
        p=0
        for name in glob.glob(path+"/"+name,recursive=True):
            #print(name)
            p1=pos("\\",name,-1)
            name3=copy(name,p1+1,len(name))
            g.getdir_files[p]=name3
            p=p+1

        return p


def IsNum(moji):
    return moji.isdigit()



def hash_data_in(hash_table, hash_key, hash_val):
    hash_table[hash_key]=hash_val
    return hash_table

def hash_data_out(hash_table, kind, num):
    #kind       "hash_key"   キー
    #           "hash_val"   値
    #           "hash_exists"   存在するか
    #num        順列番号またはexists時のkey

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
    #hash_atai      "削除するキー　""で全削除

    if hash_atai=="":
        hash_table.clear()
        return len(hash_table)

    else:
        del hash_table[hash_atai]
        return len(hash_table)


def hash_sort(hashtable,kind=""):
    '''
    hashtable:    配列
    kind:""    　　　昇順
         "reverse"   降順
    '''
    if kind=="":
        hashtable2 = dict(sorted(hashtable.items()))
        return hashtable2
    elif kind=="reverse":
        hashtable2 = dict(sorted(hashtable.items(),reverse=True))
        return hashtable2



def fopen(file,kind):
    '''
    file:ファイルのパス
    kind
        "f_exists";存在有無
        "f_read":読む
        "f_write":新規
        "f_read_or_f_write":読み書き
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
            f = open(file, 'r', encoding='utf-8-sig')

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



def fclose(fdata):
    if fdata["0_3"]=="read_only":
        print("")
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

        f = open(path, 'w', encoding='UTF-8')
        f.writelines(f_str)
        f.close()



def fdelline(fdata,gyou):
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
    screen_w,screen_h = pyautogui.size()
    g.g_screen_w=screen_w
    g.g_screen_h=screen_h


def str_conv(moji,kind):
    '''
    moji   変換する文字
    kind    0    大文字⇒小文字
            1    小文字⇒大文字
            2    全角⇒半角
            3    半角⇒全角
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
    #moji
    #num1:開始文字位置
    #num2:コピー文字数

    num2=num1+num2
    return moji[(num1-1):(num2-1)]


def replace(moji,str1,str2):
    #str1:検索する文字
    #str2:置換する文字

    return moji.replace(str1,str2)


def token(kugiri,moji):
    #kugiri:区切り文字
    #moji:元の文字=======g.token_str


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


def betweenstr(moji,mojiA,mojiB,num):

    pos1=pos(mojiA,moji,num)
    if pos1==0:
        msg("文字がありません",0)
        return False
    else:
        pos1=pos1+len(mojiA)
        mojiC=copy(moji,pos1,len(moji))
        pos2=pos(mojiB,mojiC,1)
        if pos2==0:
            msg("文字がありません",0)
            return False
        else:
            pos2=pos2+pos1-1
            getstr=copy(moji,pos1,(pos2-pos1))

            return getstr



def File_system(kind,pathA,pathB=""):
    '''
    kind
        "fs_createfolder"      :フォルダ作成
        "fs_createtextfile"        :ファイル作成
        "fs_copyfile"          :ファイル作成
        "fs_copyfolder"        :ファイル作成
        "fs_exists"            :存在確認
        "fs_deletefile"         :ファイル削除
        "fs_deletefolder"       :フォルダ削除
    '''

    import shutil

    if kind=="fs_createfolder":
        os.makedirs(pathA)
    elif kind=="fs_createtextfile":
        f = open(pathA, 'w')
        f.write('')  # 何も書き込まなくてファイルは作成されました
        f.close()
    elif kind=="fs_copyfile":
        if pathB=="":
            msg("コピー先を指定してください",0)
        else:
            shutil.copy2(pathA,pathB)
    elif kind=="fs_copyfolder":
        if pathB=="":
            msg("コピー先を指定してください",0)
        else:
            shutil.copytree(pathA,pathB)
    elif kind=="fs_exists":
        if os.path.exists(pathA)==False:
            return False
        else:
            return True
    elif kind=="fs_deletefile":
        os.remove(pathA)
    elif kind=="fs_deletefolder":
        shutil.rmtree(pathA)


def msg(msg_str,msg_title="msg",kind=0):
    '''
    msg_title:      タイトル
    msg_str:        メッセージ
    kind    :0      情報
            :1      警告
            :2      エラー
            :3      OKボタンのみ
            :4      はい・いいえ              はいボタン・・・True   いいえボタン・・・False
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
    while True:
        hwnd=getid(h1)
        if hwnd > 0:break

    ctrlwin(hwnd,"fs_topmost")



def inputbox(msg_title,msg_str,kind):

    import tkinter.simpledialog as simpledialog
    import tkinter.filedialog as filedialog

    tk.Tk().withdraw()

    # 入力ダイアログ
    if kind==0:
        ret = simpledialog.askstring(msg_title,msg_str)
    elif kind==1:
        ret = filedialog.askopenfilename()

    return ret


def gettime(g_day):
    '''
    0      :今日
    1      :1日後
    -1     :1日前
    '''

    import datetime

    g_day = g_day*-1
    if g_day==0:
        td=datetime.datetime.now()
    else:
        td=datetime.datetime.now()-datetime.timedelta(days=g_day)

    g.token_str=str(td)


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

    return td



def fukidasi(msg="",color="skyblue"):
    '''
    fukidasi(msg,color)

    msg        ；表示する文字
               ；省略でfukidasiを消去
    color      ；背景色
               white,lightyellow,skyblue,lightgreen,pink,silver,blue,aqua,yellow,orange,red,gray,fuchsia,brownなど
    '''

    import warnings

    warnings.filterwarnings('ignore')
    if msg != "":
        if getid("FUKIDASI - Google Chrome") == 0:
            IE_msg_str = "<html>"
            IE_msg_str=IE_msg_str+"<head>"
            IE_msg_str=IE_msg_str+"<title> FUKIDASI </title>"
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
            g.driver_fuki = webdriver.Chrome(options=options)
            g.driver_fuki.set_window_size(g.g_screen_w, 200)
            g.driver_fuki.get(g.CUR_DIR+"//Temp.html")
            BusyWait(g.driver_fuki)
            ###########IE_msg###########

            id = getid("FUKIDASI - Google Chrome")
            ctrlwin(id,"fs_topmost")
            ACW(id, 0, 0, g.g_screen_w, 200)
        else:

            g.driver_fuki.execute_script("document.getElementById('msg').innerText="+chr(34)+msg+chr(34))

    else:
        if getid("FUKIDASI - Google Chrome") != 0:
            g.driver_fuki.quit()

    warnings.filterwarnings('default')










def IE_msg(msg, btn_type="btn_ok",sizex=500,sizey=150,TimeOut=600):
    '''
    IE_msg(msg,[btn_type,sizex,sizey,timeout])

    msg；文字
    btn_type；
        "btn_ok"
        "btn_ok_or_btn_no"
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


    #//###########IE_msg###########
    driver0 = webdriver.Chrome(options=options)
    driver0.set_window_size(sizex + 400, sizey + 100)
    driver0.get(g.CUR_DIR+"/Temp.html")
    BusyWait(driver0)
    #//###########IE_msg###########

    getdata=driver0.execute_script("return document.getElementById('btn1').focus()")


    time.sleep(100/1000)

    try:
        msg_cyc = 0
        while True:
            getdata = driver0.execute_script("return document.getElementById('btn1').value")

            if getdata == "-":
                driver0.quit()
                return "btn_ok"

            if btn_type == "btn_ok_or_btn_no":
                getdata = driver0.execute_script("return document.getElementById('btn2').value")

                if getdata == "-":
                    driver0.quit()
                    return "btn_no"

            sleep (100)

            if ((msg_cyc / 10) > TimeOut):
                driver0.quit()
                return "timeout"

            msg_cyc = msg_cyc + 1



    except:
        driver0.quit()
        return "IE_close"


def IE_input(msg,default="",path_sen=False,password=False):

    '''
'    IE_input(msg,default,[path_sen,pasword])
'
'        msg；文字
'        default；デフォルトの入力文字
'        path_sen；
'            false；通常のinputbox
'            true；ディレクトリ選択画面にする
'        pasword；
'            true；入力文字を*で表示
'            false；入力文字は通常表示

'        戻り値；入力された文字　キャンセル時はfalse
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


        #'''//###########IE_msg###########
        driver0 = webdriver.Chrome(options=options)
        driver0.set_window_size(g.g_screen_w * 4 / 5, 400)
        #'''//###########IE_msg###########


        driver0.get(g.CUR_DIR+"//Temp.html")
        BusyWait(driver0)

        driver0.execute_script("document.getElementById('atai').focus()")

        sleep (100)

        if default != "":
            driver0.execute_script("document.getElementById('atai').Select")

        try:
            input_cyc = 0
            while True:
                getdata = driver0.execute_script("return document.getElementById('btn1').value")
                if getdata == "-":
                    driver0.quit
                    return driver0.execute_script("return document.getElementById('atai').value")

                getdata = driver0.execute_script("return document.getElementById('btn2').value")


                if getdata == "-":
                    driver0.quit
                    return False

                sleep (100)

                if (input_cyc / 10) > TimeOut:
                    driver0.quit
                    return "timeout"

                getdata = driver0.execute_script("return document.getElementById('eventAA').innerText")
                if getdata != "":
                    driver0.quit
                    return driver0.execute_script("return document.getElementById('atai').value")

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

        #'''//###########IE_msg###########
        driver0 = webdriver.Chrome(options=options)
        driver0.set_window_size(g.g_screen_w * 4/5, 400)
        #'''//###########IE_msg###########


        driver0.get(g.CUR_DIR+"//Temp.html")
        BusyWait(driver0)


        driver0.execute_script("return document.getElementById('btn1').focus()")

        sleep (100)

        try:
            input_cyc = 0
            while True:
                getdata = driver0.execute_script("return document.getElementById('btn1').value")

                if getdata == "-":
                    driver0.quit
                    return driver0.execute_script("return document.getElementById('atai').value")

                getdata = driver0.execute_script("return document.getElementById('btn2').value")

                if getdata == "-":
                    driver0.quit
                    return False

                sleep (100)

                if (input_cyc/10) > TimeOut:
                    driver0.quit
                    return "timeout"

                getdata = driver0.execute_script("return document.getElementById('eventAA').innerText")
                if getdata != "":
                    driver0.quit
                    return driver0.execute_script("return document.getElementById('atai').value")

                input_cyc = input_cyc + 1

        except:
          return "web_close"


def IE_GetTable(driver0,table_num=0,row_num=-1,cell_num=-1):
    '''
'    IE_gettable(driver0,table_num,[row_num,cell_num])
'
'        driver0；driverのオブジェクト
'        table_num；tableのitem番号
'        row_num；行
'        cell_num；列

'        戻り値；取得した文字
'           行省略時は　行数_列数 を返す

'           行数；g.IE_table_rows
'           列数；g.IE_table_cells
'           データ；g.IE_table(行数 & "_" & 列数)　に格納
    '''

    g.IE_table ={}

    driver0.execute_script("IE_tb=document.getElementsByTagName('table').item("+table_num+")")
    ie_gyou = driver0.execute_script("return IE_tb.rows.length")
    ie_retu = driver0.execute_script("return IE_tb.rows[0].cells.length")

    g.IE_table_rows = ie_gyou
    g.IE_table_cells = ie_retu

    #行省略時
    if row_num==-1:
        return str(ie_gyou)+"_"+str(ie_retu)

    #テーブルの値を配列に入れる
    for i in range(0,ie_gyou - 1):
        for ii in range(0,ie_retu - 1):
            getdata = driver0.execute_script("return IE_tb.rows["+i+"].cells["+ii+"].innerText")
            g.IE_table[str(i)+"_"+str(ii)]= getdata

    return g.IE_table[str(row_num)+"_"+str(cell_num)]




def popupmenu(selectpopup,msg= "選択してください"):
    '''
    popupmenu (selectpopup)

    selectpopup；選択するための配列

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
    driver0 = webdriver.Chrome(options=options)

    driver0.get(g.CUR_DIR+"\Temp.html")
    BusyWait(driver0)
    driver0.execute_script("return document.getElementById('sel').focus()")



    try:
        while True:
            getdata = driver0.execute_script("return document.getElementById('event').innerText")

            if getdata != "" and getdata != -1:
                if getdata == "end":
                    g.popup_result = -1
                    return -1
                else:
                    getnum = int(getdata)
                    g.popup_result = selectpopup[getnum]
                    driver0.quit()
                    return getnum

            sleep(10)

    except:

        g.popup_result = -1
        return -1


def popupmenu2(table):
    import PySimpleGUI as sg

    menu_def = None
    ldict = {}


    menu_def = "g.menu_def=[['Menu',["
    for i in range(0,len(table)):
        menu_def=menu_def+"'"+table[i]+"',"
    menu_def=menu_def+"]]]"

    #print(menu_def)
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
            getdata=0
            break

        if getflag == True:
            break

    window.close()

    Thread_end(th1)
    return getdata


def pop_click():
    while True:
        id=getid("Menuの設定","TkTopLevel")
        if id > 0:break
        sleep(10)

    ctrlwin(id,"fs_active")
    mmv(g.g_screen_w/2-50,g.g_screen_h/2+15)
    btn("left")



def popupmenu_pro(table):
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
    while True:
        getdata = driver0.execute_script("return document.readyState")
        if getdata=="complete":break
        sleep(100)



def calender_pro(kind=1):
    '''
    'kind           :0      2023-04-01
    'kind           :1      20230401
    'kind           :2      2023/04/01
    '''

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
    web_str = web_str+"<input type='date' id='date1' name='calender' onchange='eventA(this.id)' value='' /></td>\n"
    web_str = web_str+"<br><br><input id='cancel' type='button' value = 'Cancel' onClick='eventA(this.id)'>\n"
    web_str = web_str+"<span id='eventA'  style='display:none;'></span></body>\n"
    web_str = web_str+"</html>"

    f1={}
    f1 = fopen(g.CUR_DIR+"//Temp.html","f_write")
    fput(f1, web_str)
    fclose(f1)

    #'''//###########IE_msg###########
    driver0 = webdriver.Chrome(options=options)
    driver0.set_window_size(g.g_screen_w * 2 / 5, 200)
    #'''//###########IE_msg###########


    driver0.get(g.CUR_DIR+"//Temp.html")
    sleep(200)
    BusyWait(driver0)
    driver0.refresh
    sleep(200)

    driver0.execute_script("document.getElementById('date1').focus()")

    try:
        while True:

            getdata = driver0.execute_script("return document.getElementById('eventA').innerText")

            if getdata == "date1":
                getdata2  = driver0.execute_script("return document.getElementById('date1').value")

                driver0.quit

                if kind == 0:
                    return getdata2
                elif kind == 1:
                    return replace(getdata2,"-","")
                elif kind == 2:
                    return replace(getdata2,"-","/")

            elif getdata == "cancel":
                driver0.quit
                return "cancel"

    except:
        return "cancel"



def exec_apli(app,wait=0):
    '''
    app    起動するアプリのパス
    wait   0   :待たない
           1　 :待つ

    j.exec_apli(r"\\tr\dfs\ELデバイス部\07_室課\03_プロ開室\神川\情報機器持出返却管理表\Temp.html")
    '''

    import subprocess

    if wait==0:
        subprocess.run(app,shell=True)
    else:
        subprocess.run(app,shell=False)


def open_folder_dialog(dir_path=""):
    '''
    dir_path     開くフォルダパス

    '''
    from tkinter import filedialog

    dir = dir_path
    fld = filedialog.askdirectory(initialdir = dir)
    return fld



def open_file_dialog(dir_path=""):
    '''
    dir_path     開くフォルダパス
    複数の時は"_"で結合

    '''

    from tkinter import filedialog

    typ = [('', '*')]
    dir = dir_path
    fle = filedialog.askopenfilenames(filetypes = typ, initialdir = dir)

    file_name=""
    for f in fle:
        if f==0:
            file_name=f
        else:
            file_name=file_name+"_"+f

    return file_name



#ゼロ埋めする
def formatA(num,keta):
    '''
    num:      元の数字
    keta:     ゼロ埋めして整える桁数
    '''
    getdata=format(num,"0"+str(keta))
    return getdata





def btn(kind,wheel=1):
    '''
    kind    "left"
            "right"
            "middle"
            "wheel"   ホイール
    wheel    ホイール操作　1:上　　-1:した
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


def mouse_position():
    '''
    マウスのポジションを得る
    ポジションは
    x位置　g.mouse_pos[0]
    y位置  g.mouse_pos[1]
    '''
    import mouse
    g.mouse_pos={}
    g.mouse_pos=mouse.get_position()
    #print(g.mouse_pos[0],g.mouse_pos[1])
    return g.mouse_pos



def mmv(posx,posy,kind="",posx2=0,posy2=0):
    '''
    posx:      移動するX位置
    posy:      移動するＹ位置
    kind:      ""        移動
               "drag"    ドラッグ
    posx2:      ドラッグするX位置
    posy2:      ドラッグするＹ位置
    '''

    import mouse

    if kind=="":
        mouse.move(posx,posy, absolute=True, duration=0.1)
    elif kind=="drag":
        mouse.drag(posx,posy,posx2,posy2,absolute = True,duration = 0.1)


def POFF(kind):
    '''
    kind:"shutdown"
        :"logoff"
    '''
    import os

    if kind=="shutdown":
        os.system("shutdown /s /t 1")
    elif kind=="logoff":
        os.system("shutdown /l")



def mailsend(kenmei,Soushin,Bunsyo,kind=0,SoushinCC=""):#件名,送信先,文書
    import win32com.client

    outlook = win32com.client.Dispatch("Outlook.Application")

    mail = outlook.CreateItem(0)

    mail.to = Soushin
    mail.cc = SoushinCC
    mail.bcc = ''
    mail.subject = kenmei
    mail.bodyFormat = 1
    mail.body = Bunsyo

    if kind==0:
        mail.display(True)
    elif kind==1:
        mail.Send()


def up_dir(path):
    '''
    path:パス
    戻り値:一つ上のディレクトリ
    '''
    import os
    return os.path.dirname(path)


def path_join(*path_str):
    '''
    *path_str:　　　　　path1,path2,path3,…………


    '''
    import os

    path_table=list()
    path_table=path_str
    return os.path.join(*path_table)



def End():
    import sys

    sys.exit()


def baloon(msg,title="通知"):

    from plyer import notification
    notification.notify(title=msg,message=title,app_name="python",)



def ex_OPEN(path=""):
    '''
    ex,wb=ex_OPEN(path)

    戻り値：   excel:アプリケーションOBJ
               wb:ワークブックOBJ


    '''
    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    if path != "":
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
    wb:         workbookオブジェクト
    ex_kind:    "Close"  保存しない
    　　　　　　"Save"   上書き保存する
                "保存先"  名前を付けて保存する
    kind2       "Quit"   excelを終わる

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



def Thread(sub,h1=0):
    '''
    shread名=Thread(関数名)
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



'''
def getid(title):
    import array
    import pywin32
    import win32con

    # 親ウィンドウハンドル(識別番号)とそのクラス名、ハンドル名を取得
    p_handle = win32gui.FindWindow(None,title) # アプリケーションのタイトル文字列など入力
    p_class_name = win32gui.GetClassName(p_handle)
    p_handle_name = win32gui.GetWindowText(p_handle)
    handles_dict = {str(p_handle): [p_class_name,p_handle_name]}
    print('p',hex(p_handle),p_class_name,p_handle_name)
'''

if __name__ == '__main__':
    main()
