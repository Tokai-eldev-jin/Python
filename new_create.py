#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      A0044
#
# Created:     21/04/2023
# Copyright:   (c) A0044 2023
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    pass

def new_create(loginUser,name_sel,name_table,name_table2,Shouninsya_sel,Shouninsya_table,Shouninsya_table2,driver):

    import jin as j
    import global_value as g


    today_YY4:str
    today_mm2:str
    today_dd2:str
    today_hh2:str
    today_nn2:str
    today_ss2:str

    i:int
    t:int
    kazu:int
    Mochidasi_list:str
    Menu_list:str
    name_table_A={}
    id:int



    ########### Chrome 起動　###########
    driver.maximize_window()
    driver.get(j.path_join(g.CUR_DIR,r"ini\index.html"))
    j.sleep (200)
    j.BusyWait(driver)
    driver.refresh()
    ########### Chrome 起動　###########



    ########### Chrome 初期入力　###########
    j.gettime(0)
    today_YY4 = g.G_TIME_YY4
    today_mm2 = g.G_TIME_MM2
    today_dd2 = g.G_TIME_DD2
    today_hh2 = g.G_TIME_HH2
    today_nn2 = g.G_TIME_NN2
    today_ss2 = g.G_TIME_SS2


    #持出日
    driver.execute_script("document.getElementById('date1').value=" + chr(34) + today_YY4 + "-" + today_mm2 + "-" + today_dd2 + chr(34))

    #持出時刻
    if len(today_hh2) == 1 : today_hh2 = "0" + today_hh2
    if len(today_nn2) == 1 : today_nn2 = "0" + today_nn2

    driver.execute_script("document.getElementById('time1').value=" + chr(34) + today_hh2 + ":" + today_nn2 + chr(34))

    #持出者
    driver.execute_script("document.getElementById('mochidashi_sya').innerHTML=" + chr(34) + name_sel + chr(34))
    driver.execute_script("document.getElementById('name_list1').value=" + chr(34) + loginUser + chr(34))

    #持出機器
    f1={}
    f1 = j.fopen(j.path_join(g.CUR_DIR,r"ini\List.txt"), "f_read")
    kazu = j.fget(f1, -1)

    Mochidasi_list = "<select id='mochidashi_list'>"
    for i  in range(1,kazu+1):
        Mochidasi_list = Mochidasi_list + "<option value='" + j.fget(f1, i, 1) + "'>" + j.fget(f1, i, 1) + "</option>"

    Mochidasi_list = Mochidasi_list + "</select>"

    driver.execute_script("document.getElementById('mochidashi_kiki').innerHTML=" + chr(34) + Mochidasi_list + chr(34))


    #返却予定日
    driver.execute_script("document.getElementById('date2').value=" + chr(34) + today_YY4 + "-" + today_mm2 + "-" + today_dd2 + chr(34))

    #返却予定時間
    j.gettime(0)
    driver.execute_script("document.getElementById('time2').value=" + chr(34) + today_hh2 + ":" + today_nn2 + chr(34))

    #確認者
    driver.execute_script("document.getElementById('Kakuninnsya').innerHTML=" + chr(34) + Shouninsya_sel + "　<input type='text' size='10' style='font-size:8;' value='' id='name_text' onchange='eventA(this.id)'>" + chr(34))
    driver.execute_script("document.getElementById('Shouninsya_list1').value=" + chr(34) + "" + chr(34))

    #disabled
    #driver.execute_script ("return document.getElementById('Soushin1').disabled=true")
    driver.execute_script ("return document.getElementById('Soushin2').disabled=true")
    driver.execute_script ("return document.getElementById('Shounin1').disabled=true")
    driver.execute_script ("return document.getElementById('Shounin2').disabled=true")
    driver.execute_script ("return document.getElementById('Sashimodoshi1').disabled=true")
    driver.execute_script ("return document.getElementById('Sashimodoshi2').disabled=true")

    driver.execute_script ("return document.getElementById('date3').disabled=true")
    driver.execute_script ("return document.getElementById('time3').disabled=true")
    driver.execute_script ("return document.getElementById('virus Check_box2').disabled=true")
    driver.execute_script ("return document.getElementById('Data Check_box').disabled=true")
    driver.execute_script ("return document.getElementById('Chuushi').disabled=true")

    driver.execute_script ("document.getElementById('syanai').click()")

    id = j.getid("情報機器持出返却管理表 - Google Chrome")




##    try:

    while True:

        getdata = driver.execute_script("return document.getElementById('event').innerText")

        if getdata != "" :
            driver.execute_script ("document.getElementById('event').innerText=" + chr(34) + "" + chr(34))

            #ラジオボタン処理
            if getdata == "syanai" :
                driver.execute_script ("return document.getElementById('kaigi').disabled = false")
                driver.execute_script ("return document.getElementById('hokan').disabled = false")
                driver.execute_script ("return document.getElementById('data_move').disabled = false")

                driver.execute_script ("return document.getElementById('kyakusaki').checked = false")
                driver.execute_script ("return document.getElementById('kyakusaki').disabled = true")
                driver.execute_script ("return document.getElementById('siiresaki').checked = false")
                driver.execute_script ("return document.getElementById('siiresaki').disabled = true")
                driver.execute_script ("return document.getElementById('sonota0').checked = false")
                driver.execute_script ("return document.getElementById('sonota0').disabled = true")
                driver.execute_script ("document.getElementById('in_sonota1').value =" + chr(34) + "" + chr(34))
                driver.execute_script ("return document.getElementById('in_sonota1').disabled = true")

                driver.execute_script ("document.getElementById('sonota1').value =" + chr(34) + "" + chr(34))
                driver.execute_script ("return document.getElementById('sonota1').disabled = true")

            elif getdata == "shagai" :
                driver.execute_script ("return document.getElementById('kaigi').checked=false")
                driver.execute_script ("return document.getElementById('kaigi').disabled = true")
                driver.execute_script ("return document.getElementById('hokan').checked=false")
                driver.execute_script ("return document.getElementById('hokan').disabled = true")
                driver.execute_script ("return document.getElementById('data_move').checked=false")
                driver.execute_script ("return document.getElementById('data_move').disabled = true")

                driver.execute_script ("return document.getElementById('kyakusaki').disabled = false")
                driver.execute_script ("return document.getElementById('siiresaki').disabled = false")
                driver.execute_script ("return document.getElementById('sonota0').disabled = false")
                driver.execute_script ("return document.getElementById('in_sonota1').disabled = false")

                driver.execute_script ("document.getElementById('sonota1').value =" + chr(34) + "" + chr(34))
                driver.execute_script ("return document.getElementById('sonota1').disabled = true")

            elif getdata == "sonota" :
                driver.execute_script ("return document.getElementById('kaigi').checked=false")
                driver.execute_script ("return document.getElementById('kaigi').disabled = true")
                driver.execute_script ("return document.getElementById('hokan').checked=false")
                driver.execute_script ("return document.getElementById('hokan').disabled = true")
                driver.execute_script ("return document.getElementById('data_move').checked=false")
                driver.execute_script ("return document.getElementById('data_move').disabled = true")

                driver.execute_script ("return document.getElementById('kyakusaki').checked = false")
                driver.execute_script ("return document.getElementById('kyakusaki').disabled = true")
                driver.execute_script ("return document.getElementById('siiresaki').checked = false")
                driver.execute_script ("return document.getElementById('siiresaki').disabled = true")
                driver.execute_script ("return document.getElementById('sonota0').checked = false")
                driver.execute_script ("return document.getElementById('sonota0').disabled = true")
                driver.execute_script ("document.getElementById('in_sonota1').value =" + chr(34) + "" + chr(34))
                driver.execute_script ("return document.getElementById('in_sonota1').disabled = true")

                driver.execute_script ("return document.getElementById('sonota1').disabled = false")



            #確認者　入力変化
            if getdata == "name_text" :

                getdata = driver.execute_script("return document.getElementById('name_text').value")
                if getdata == "" : break

                name_table_A = {}
                t = 0
                for i  in range(0,len(Shouninsya_table)):
                    if j.pos(getdata, j.hash_data_out(Shouninsya_table, "hash_val", i)) > 0 :
                        name_table_A[t] = j.hash_data_out(Shouninsya_table, "hash_val", i)
                        t = t + 1

                if t == 0 :
                    pass
                elif t == 1 :
                    getdata = name_table_A[0]
                    getdata = Shouninsya_table2[getdata]
                    driver.execute_script ("document.getElementById('Shouninsya_list1').value=" + chr(34) + getdata + chr(34))
                else:
                    getdata = j.popupmenu2(name_table_A)

                    if getdata == -1 :
                        pass
                    else:
                        getdata = name_table_A[getdata]
                        getdata = Shouninsya_table2[getdata]
                        driver.execute_script ("return document.getElementById('Shouninsya_list1').value="+chr(34)+getdata+chr(34))
            #終了
            if getdata == "syuuryou" :
                return ""

            #送信
            if getdata == "Soushin1" :
                持出日:str =""
                持出時刻:str =""
                持出者:str =""
                持出機器:str =""
                撮影物:str =""
                撮影場所:str =""
                社内:bool =0
                会議資料 :bool =0
                保管用 :bool =0
                データ移動 :bool =0
                社外 :bool =0
                客先 :bool =0
                仕入先 :bool =0
                その他0 :bool =0
                INその他1:str =""
                その他 :bool =0
                その他1:str =""
                返却予定日:str =""
                返却予定時間:str =""
                ウィルスチェック :bool =0
                確認者:str =""
                返却日:str =""
                返却時間:str =""
                ウィルスチェック2 :bool =0
                データ削除 :bool =0
                確認者2:str =""
                file_id:str =""
                status_num:int =0

                koumoku_list = {}
                koumoku_list["持出日"] = 持出日                         #0
                koumoku_list["持出時刻"] = 持出時刻                     #1
                koumoku_list["持出者"] = 持出者                         #2
                koumoku_list["持出機器"] = 持出機器                     #3
                koumoku_list["撮影物"] = 撮影物                         #4
                koumoku_list["撮影場所"] = 撮影場所                     #5
                koumoku_list["社内"] = 社内                             #6
                koumoku_list["会議資料"] = 会議資料                     #7
                koumoku_list["保管用"] = 保管用                         #8
                koumoku_list["データ移動"] = データ移動                     #9
                koumoku_list["社外"] = 社外                             #10
                koumoku_list["客先"] = 客先                             #11
                koumoku_list["仕入先"] = 仕入先                         #12
                koumoku_list["その他0"] = その他0                       #13
                koumoku_list["INその他1"] = INその他1                   #14
                koumoku_list["その他"] = その他                         #15
                koumoku_list["その他1"] = その他1                       #16
                koumoku_list["返却予定日"] = 返却予定日                 #17
                koumoku_list["返却予定時間"] = 返却予定時間             #18
                koumoku_list["ウィルスチェック"] = ウィルスチェック                 #19
                koumoku_list["確認者"] = 確認者                         #20
                koumoku_list["返却日"] = 返却日                         #21
                koumoku_list["返却時間"] = 返却日                       #22
                koumoku_list["ウィルスチェック2"] = ウィルスチェック2               #23
                koumoku_list["データ削除"] = データ削除                   #24
                koumoku_list["確認者2"] = 確認者2                       #25


                #入力チェック
                持出日 = driver.execute_script("return document.getElementById('date1').value")
                koumoku_list["持出日"] = 持出日
                if 持出日 == "" :
                    j.msg("持出日を入力してください")
                    continue




                持出時刻 = driver.execute_script("return document.getElementById('time1').value")
                koumoku_list["持出時刻"] = 持出時刻
                if 持出時刻 == "" :
                    j.msg("持出時刻を入力してください")
                    continue

                持出者 = driver.execute_script("return document.getElementById('name_list1').value")
                koumoku_list["持出者"] = 持出者
                if 持出者 == "" :
                    j.msg("持出者を入力してください")
                    continue

                持出機器 = driver.execute_script("return document.getElementById('mochidashi_list').value")
                koumoku_list["持出機器"] = 持出機器
                if 持出機器 == "" :
                    j.msg("持出機器を入力してください")
                    continue

                撮影物 = driver.execute_script("return document.getElementById('satuei_butu_box').value")
                koumoku_list["撮影物"] = 撮影物
                ##                    if 撮影物 = "" And pos("カメ", 持出機器) > 0 :
                ##                    baloon id, "撮影物を入力してください",  "message",0)
                ##                    GoTo continue1
                ##


                撮影場所 = driver.execute_script("return document.getElementById('satuei_basho_box').value")
                koumoku_list["撮影場所"] = 撮影場所
                ##    '                if 撮影場所 == "" And pos("カメ", 持出機器) > 0 :
                ##    '                    baloon id, "撮影場所を入力してください",  "message",0)
                ##    '                    GoTo continue1
                ##    '


                社内 = driver.execute_script("return document.getElementById('syanai').checked")
                koumoku_list["社内"] = 社内
                if 社内 == True :
                    会議資料 = driver.execute_script("return document.getElementById('kaigi').checked")
                    保管用 = driver.execute_script("return document.getElementById('hokan').checked")
                    データ移動 = driver.execute_script("return document.getElementById('data_move').checked")

                    koumoku_list["会議資料"] = 会議資料
                    koumoku_list["保管用"] = 保管用
                    koumoku_list["データ移動"] = データ移動

                    if 会議資料 == False and 保管用 == False and データ移動 == False :
                        j.msg("社内のチェックボックスを入力してください")
                        continue

                else:
                    社外 = driver.execute_script("return document.getElementById('shagai').checked")
                    koumoku_list["社外"] = 社外
                    if 社外 == True :
                        客先 = driver.execute_script("return document.getElementById('kyakusaki').checked")
                        仕入先 = driver.execute_script("return document.getElementById('siiresaki').checked")
                        その他0 = driver.execute_script("return document.getElementById('sonota0').checked")
                        koumoku_list["客先"] = 客先
                        koumoku_list["仕入先"] = 仕入先
                        koumoku_list["その他0"] = その他0

                        if 客先 == False and (仕入先 == False and その他0 == False) :
                            j.msg("社外のチェックボックスを入力してください")
                            continue

                        if その他0 == True :
                            INその他1 = driver.execute_script("return document.getElementById('in_sonota1').value")
                            koumoku_list["INその他1"] = INその他1
                            if INその他1 == "" :
                                j.msg("その他を入力してください")
                                continue
                    else:
                        その他 = driver.execute_script("return document.getElementById('sonota').checked")
                        koumoku_list["その他"] = その他
                        if その他 == True :
                            その他1 = driver.execute_script("return document.getElementById('sonota1').value")
                            koumoku_list["その他1"] = その他1
                            if その他1 == "" :
                                j.msg("その他を入力してください")
                                continue

                        else:
                            j.msg("使用目的を入力してください")
                            continue


                返却予定日 = driver.execute_script("return document.getElementById('date2').value")
                koumoku_list["返却予定日"] = 返却予定日
                if 返却予定日 == "" :
                    j.msg("返却予定日を入力してください")
                    continue

                返却予定時間 = driver.execute_script("return document.getElementById('time2').value")
                koumoku_list["返却予定時間"] = 返却予定時間
                if 返却予定時間 == "" :
                    j.msg("返却予定時間を入力してください")
                    continue

                ウィルスチェック = driver.execute_script("return document.getElementById('virus Check_box').checked")
                koumoku_list["ウィルスチェック"] = ウィルスチェック
                if ウィルスチェック == False :
                    j.msg("ウィルスチェックをしてください")
                    continue

                確認者 = driver.execute_script("return document.getElementById('Shouninsya_list1').value")
                koumoku_list["確認者"] = 確認者
                if 確認者 == "" :
                    j.msg("確認者を入力してください")
                    continue

                if 確認者 == 持出者 :
                    j.msg("持出者と確認者が同じです。")
                    continue

                返却日 = driver.execute_script("return document.getElementById('date3').value")
                koumoku_list["返却日"] = 返却日

                返却時間 = driver.execute_script("return document.getElementById('time3').value")
                koumoku_list["返却時間"] = 返却時間

                ウィルスチェック2 = driver.execute_script("return document.getElementById('virus Check_box2').checked")
                koumoku_list["ウィルスチェック2"] = ウィルスチェック2

                データ削除 = driver.execute_script("return document.getElementById('Data Check_box').checked")
                koumoku_list["データ削除"] = データ削除

                確認者2 = ""
                koumoku_list["確認者2"]= 確認者2



                #確認者フォルダの作成
                if j.fopen(j.path_join(g.CUR_DIR,"user",確認者), "f_exists") == False :
                    j.File_system("fs_createfolder", j.path_join(g.CUR_DIR,"user",確認者))

                status_num = 1

                #ユーザーファイル、データファイルに保存

                #datafile
                if j.fopen(j.path_join(g.CUR_DIR,"ini",today_YY4+"_datafile.txt"), "f_exists") == False :
                    f1 = j.fopen(j.path_join(g.CUR_DIR,"ini",today_YY4 + "_datafile.txt"), "f_write")
                    j.fput(f1, today_YY4, 1, 1)
                    j.fput(f1, 1, 1, 2)
                    j.fclose(f1)

                else:
                    pass

                #datafileに書き込み
                while True:

                    #ダミーファイルがない時
                    if j.fopen(j.path_join(g.CUR_DIR,"ini","dami_"+today_YY4+"_datafile.txt"),"f_exists") == False :
                        #ダミーファイルの作成
                        j.File_system("fs_createtextfile",j.path_join(g.CUR_DIR,"ini","dami_"+today_YY4+"_datafile.txt"))
                        print(j.path_join(g.CUR_DIR,"ini","dami_"+today_YY4+"_datafile.txt"))

                        f1 = j.fopen(j.path_join(g.CUR_DIR,"ini",today_YY4+"_datafile.txt"), "f_read_or_f_write")
                        kazu = j.fget(f1, -1)
                        file_id = j.fget(f1, kazu, 1) + "_" + j.fget(f1, kazu, 2)

                        #status
                        j.fput(f1, status_num, kazu, 3)     #status   1:作成時　2：持出サイン時　3:返却送信時　4：完了時  5:差戻1　　6:差戻2

                        #data
                        for i  in range(0,len(koumoku_list)):
                            j.fput(f1, j.hash_data_out(koumoku_list,"hash_val", i), kazu, 4 + i)

                        #next saiban
                        j.fput(f1, j.fget(f1, kazu, 1), kazu + 1, 1)
                        j.fput(f1, int(j.fget(f1, kazu, 2)) + 1, kazu + 1, 2)
                        j.fclose(f1)

                        j.File_system("fs_deletefile",j.path_join(g.CUR_DIR ,"ini","dami_"+today_YY4+"_datafile.txt"))

                        break
                    else:
                        pass


                #USER_fileに書き込み
                while True:
                    if j.fopen(j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"),"f_exists") == False :
                        #dami
                        j.File_system("fs_createtextfile",j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"))

                        f1 = j.fopen(j.path_join(g.CUR_DIR,"user",持出者,持出者+".txt"),"f_read_or_f_write")
                        kazu = j.fget(f1, -1)

                        if kazu > 1000 :
                            j.fdelline(f1, 1)
                            j.fclose(f1)

                            f1 = j.fopen(j.path_join(g.CUR_DIR,"user",持出者,持出者+".txt"), "f_read_or_f_write")
                            kazu = j.fget(f1, -1)

                        #採番
                        g.token_str = file_id
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                        #status
                        j.fput(f1, status_num, kazu + 1, 3) #status   1:作成時　2：持出サイン時　3:返却送信時　4：完了時  5:差戻1　　6:差戻2

                        #data
                        for i  in range(0,len(koumoku_list)):
                            j.fput(f1, j.hash_data_out(koumoku_list,"hash_val", i), kazu + 1, 4 + i)

                        j.fclose(f1)

                        j.File_system("fs_deletefile",j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"))

                        break
                    else:
                        pass


                #USER_Re_fileに書き込み
                while True:
                    if j.fopen(j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"),"f_exists") == False :
                        #dami
                        j.File_system("fs_createtextfile",j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"))

                        f1 = j.fopen(j.path_join(g.CUR_DIR,"user",持出者,"Re_"+持出者+".txt"),"f_read_or_f_write")
                        kazu = j.fget(f1, -1)

                        if kazu > 1000 :
                            j.fdelline(f1, 1)
                            j.fclose(f1)

                            f1 = j.fopen(j.path_join(g.CUR_DIR,"user",持出者,"\Re_"+持出者+".txt"), "f_read_or_f_write")
                            kazu = j.fget(f1, -1)


                        #採番
                        g.token_str = file_id
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                        #status
                        j.fput(f1, 0, kazu + 1, 3)  #'status   0:無効　1:作成時　2：持出サイン時　3:返却送信時　4：完了時  5:差戻1　　6:差戻2

                        j.fclose(f1)

                        j.File_system("fs_deletefile",j.path_join(g.CUR_DIR,"user",持出者,"dami_"+持出者+".txt"))

                        break
                    else:
                        pass

                #確認者_Re_fileに書き込み
                while True:
                    if j.fopen(j.path_join(g.CUR_DIR,"user",確認者,"dami_"+確認者+".txt"),"f_exists") == False :
                        #dami
                        j.File_system("fs_createtextfile",j.path_join(g.CUR_DIR,"user",確認者,"dami_"+確認者+".txt"))

                        f1 = j.fopen(j.path_join(g.CUR_DIR,"user",確認者,"Re_"+確認者+".txt"), "f_read_or_f_write")
                        kazu = j.fget(f1, -1)

                        if kazu > 1000 :
                            j.fdelline(f1, 1)
                            j.fclose(f1)

                            f1 = j.fopen(j.path_join(g.CUR_DIR,"user",確認者,"Re_"+確認者+".txt"), "f_read_or_f_write")
                            kazu = j.fget(f1, -1)

                        #saiban
                        g.token_str = file_id
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 1)
                        j.fput(f1, j.token("_", g.token_str), kazu + 1, 2)

                        #status
                        j.fput(f1, status_num, kazu + 1, 3)

                        j.fclose(f1)

                        j.File_system("fs_deletefile",j.path_join(g.CUR_DIR,"user",確認者,"dami_"+確認者+".txt"))

                        break
                    else:
                        pass

                件名:str
                送信先:str
                送信先CC:str
                文書:str
                send:int
                exe_path:str

                exe_path = j.path_join(g.CUR_DIR,"\情報機器持出管理.vbs")

                件名 = "【申請】情報機器持出申請 " + file_id
                送信先 = koumoku_list["確認者"]

                if koumoku_list["持出者"] == loginUser :
                    送信先CC = koumoku_list["持出者"]
                else:
                    送信先CC = koumoku_list["持出者"] + "," + loginUser


                文書 = "情報機器持出申請です。\n\n"
                文書 = 文書 + "No" + file_id + "\n"
                文書 = 文書 + "以下より確認をお願いします｡ \n\n"
                文書 = 文書 + "<" + exe_path + ">"


                send = 2

                ##'                sendmail 件名, 送信先, 送信先CC, 文書, send
                ##'                driver.quit()
                ##'                driver = Nothing

                ##j.msg("申請メールを送信しました。  No." + file_id)

                return 確認者

        j.sleep (10)
##    except:
##        pass

    #continue1:
    #finish:
    driver.quit()

    return ""








if __name__ == '__main__':
    main()
