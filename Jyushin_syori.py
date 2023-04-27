#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      A0044
#
# Created:     24/04/2023
# Copyright:   (c) A0044 2023
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    pass


def Jyushin_syori(loginUser,name_sel,name_table,name_table2,Shouninsya_sel,Shouninsya_table,Shouninsya_table2,driver):
    import jin as j
    import global_value as g


    import File_write as p1

    kazu:int
    i:int
    t:int
    file_id:str
    koumoku_list:str
    do_gyou:int
    Mochidasi_list:str
    Komento:str
    do_year:str
    today_YY4:str
    today_mm2:str
    today_dd2:str
    today_hh2:str
    today_nn2:str
    today_ss2:str
    status_num:int

    件名:str
    送信先:str
    送信先CC:str
    文書:str
    send:int
    exe_path:str

##continue0:
    GoTo_continue0=0

    while True:
        GoTo_continue0=0

        j.gettime(0)
        today_YY4 = g.G_TIME_YY4
        today_mm2 = g.G_TIME_MM2
        today_dd2 = g.G_TIME_DD2
        today_hh2 = g.G_TIME_HH2
        today_nn2 = g.G_TIME_NN2
        today_ss2 = g.G_TIME_SS2
        if len(today_hh2) == 1 : today_hh2 = "0"+today_hh2
        if len(today_nn2) == 1 : today_nn2 = "0"+today_nn2

        exe_path = j.path_join(g.CUR_DIR+r"情報機器持出管理.vbs")


        ##Re_fileで未実施項目を抽出する
        Item_list = {}
        t = 0

        f1 = j.fopen(j.path_join(g.CUR_DIR,"user",loginUser,"Re_"+loginUser+".txt"), "f_read")
        kazu = j.fget(f1, -1)

        for i in range(1,kazu+1):
            status_num = int(j.fget(f1, i, 3))

            if status_num != 0 and status_num != 4 :
                Item_list[t] = j.fget(f1, i, 1)+"_"+j.fget(f1, i, 2)
                t = t + 1
            else:
                pass

        j.fclose(f1)


        if len(Item_list) == 0 : return ""


        ##選択メニュ
        if len(Item_list) > 1 :
            getdata = int(j.popupmenu(Item_list, name_table[loginUser]+"さん。<br>受信ファイルを選択してください"))
        else:
            getdata = 0


        if getdata == -1 : return ""
        file_id = Item_list[getdata]
        do_year = j.copy(file_id, 1, 4)


        ##データ抽出
        koumoku_list = {}

        f1 = j.fopen(j.path_join(g.CUR_DIR,"ini",do_year+"_datafile.txt"), "f_read")
        kazu = j.fget(f1, -1)

        for  i  in range(1,kazu+1):
            getdata = j.fget(f1, i, 1)+"_"+j.fget(f1, i, 2)

            if getdata == file_id :
                do_gyou = i

                status_num = int(j.fget(f1, do_gyou, 3)) ##status   1:作成時　2：持出サイン時　3:返却送信時　4：完了時  5:差戻1　　6:差戻2

                for t in range(0,26):
                    koumoku_list[t] = j.fget(f1, do_gyou, 4 + t)

                break

        j.fclose(f1)



        ########### Chrome 起動　###########
        ##driver As New Selenium.WebDriver
        ##driver.SetCapability "goog:chromeOptions", "{""excludeSwitches"":[""enable-automation""]}"
        ##driver.Start "chrome"
        driver.maximize_window()
        driver.get(j.path_join(g.CUR_DIR,"ini","index.html"))
        j.sleep (200)
        j.BusyWait(driver)
        driver.refresh()
        j.sleep (200)
        ########### Chrome 起動　###########


        if status_num != 5 :

            ##持出日
            driver.execute_script("document.getElementById('Mochidashi_bi').innerText="+chr(34)+koumoku_list[0]+chr(34))

            ##持出時刻
            driver.execute_script("document.getElementById('Mochidashi_jikoku').innerText="+chr(34)+koumoku_list[1]+chr(34))

            ##持出者
            driver.execute_script("document.getElementById('mochidashi_sya').innerText="+chr(34)+name_table[koumoku_list[2]]+chr(34))

            ##持出機器
            driver.execute_script("document.getElementById('mochidashi_kiki').innerText="+chr(34)+koumoku_list[3]+chr(34))

            ##撮影物
            driver.execute_script("document.getElementById('satuei_butu_box').value="+chr(34)+koumoku_list[4]+chr(34))
            driver.execute_script("return document.getElementById('satuei_butu_box').disabled=true")

            ##撮影場所
            driver.execute_script("document.getElementById('satuei_basho_box').value="+chr(34)+koumoku_list[5]+chr(34))
            driver.execute_script("return document.getElementById('satuei_basho_box').disabled=true")

            ##社内
            if koumoku_list[6] == "True" :
                driver.execute_script("return document.getElementById('syanai').checked = true")
            else:
                driver.execute_script("return document.getElementById('syanai').checked = false")

            driver.execute_script("return document.getElementById('syanai').disabled = true")

            ##会議資料
            if koumoku_list[7] == "True" :
                driver.execute_script("return document.getElementById('kaigi').checked = true")
            else:
                driver.execute_script("return document.getElementById('kaigi').checked = false")

            driver.execute_script("return document.getElementById('kaigi').disabled = true")

            ##保管用
            if koumoku_list[8] == "True" :
                driver.execute_script("return document.getElementById('hokan').checked = true")
            else:
                driver.execute_script("return document.getElementById('hokan').checked = false")

            driver.execute_script("return document.getElementById('hokan').disabled = true")

            ##データ移動
            if koumoku_list[9] == "True" :
                driver.execute_script("return document.getElementById('data_move').checked = true")
            else:
                driver.execute_script("return document.getElementById('data_move').checked = false")

            driver.execute_script("return document.getElementById('data_move').disabled = true")

            ##社外
            if koumoku_list[10] == "True" :
                driver.execute_script("return document.getElementById('shagai').checked = true")
            else:
                driver.execute_script("return document.getElementById('shagai').checked = false")

            driver.execute_script("return document.getElementById('shagai').disabled = true")

            ##客先
            if koumoku_list[11] == "True" :
                driver.execute_script("return document.getElementById('kyakusaki').checked = true")
            else:
                driver.execute_script("return document.getElementById('kyakusaki').checked = false")

            driver.execute_script("return document.getElementById('kyakusaki').disabled = true")

            ##仕入先
            if koumoku_list[12] == "True" :
                driver.execute_script("return document.getElementById('siiresaki').checked = true")
            else:
                driver.execute_script("return document.getElementById('siiresaki').checked = false")

            driver.execute_script("return document.getElementById('siiresaki').disabled = true")

            ##その他0
            if koumoku_list[13] == "True" :
                driver.execute_script("return document.getElementById('sonota0').checked = true")
            else:
                driver.execute_script("return document.getElementById('sonota0').checked = false")

            driver.execute_script("return document.getElementById('sonota0').disabled = true")

            ##INその他1
            driver.execute_script("return document.getElementById('in_sonota1').value = "+chr(34)+koumoku_list[14]+chr(34))
            driver.execute_script("return document.getElementById('in_sonota1').disabled = true")

            ##その他
            if koumoku_list[15] == "True" :
                driver.execute_script("return document.getElementById('sonota').checked = true")
            else:
                driver.execute_script("return document.getElementById('sonota').checked = false")

            driver.execute_script("return document.getElementById('sonota').disabled = true")

            ##その他1
            driver.execute_script("document.getElementById('sonota1').innerText="+chr(34)+koumoku_list[16]+chr(34))
            driver.execute_script("document.getElementById('sonota1').disabled=true")

            ##返却予定日
            driver.execute_script("document.getElementById('henkyaku_yoteibi').innerText="+chr(34)+koumoku_list[17]+chr(34))

            ##返却予定時間
            driver.execute_script("document.getElementById('henkyaku_yoteijikan').innerText="+chr(34)+koumoku_list[18]+chr(34))

            ##ウィルスチェック
            if koumoku_list[19] == "True" :
                driver.execute_script("return document.getElementById('virus Check_box').checked = true")
            else:
                driver.execute_script("return document.getElementById('virus Check_box').checked = false")

            driver.execute_script("return document.getElementById('virus Check_box').disabled = true")

            ##確認者
            if status_num == 1 :
                driver.execute_script("document.getElementById('Kakuninnsya').innerHTML="+chr(34)+"<font size='5' color='red'>&#12958;</font>　"+name_table[koumoku_list[20]]+chr(34))
            elif status_num == 2 or (status_num == 3 or status_num == 6) :
                driver.execute_script("document.getElementById('Kakuninnsya').innerHTML="+chr(34)+"<font size='5' color='gray'>&#12958;</font>　"+name_table[koumoku_list[20]]+chr(34))




        else: ##status_num=5

            ##持出日
            driver.execute_script("document.getElementById('date1').value="+chr(34)+koumoku_list[0]+chr(34))

            ##持出時刻
            driver.execute_script("document.getElementById('time1').value="+chr(34)+koumoku_list[1]+chr(34))

            ##持出者
            driver.execute_script("document.getElementById('mochidashi_sya').innerHTML="+chr(34)+name_sel+chr(34))
            driver.execute_script("document.getElementById('name_list1').value="+chr(34)+koumoku_list[2]+chr(34))

            ##持出機器
            f1 = j.fopen(j.path_join(g.CUR_DIR,"ini","List.txt"), "f_read")
            kazu = j.fget(f1, -1)

            Mochidasi_list = "<select id='mochidashi_list'>"
            for  i  in range(1,kazu+1):
                Mochidasi_list = Mochidasi_list+"<option value='"+j.fget(f1, i, 1)+"'>"+j.fget(f1, i, 1)+"</option>"

            Mochidasi_list = Mochidasi_list+"</select>"

            driver.execute_script("document.getElementById('mochidashi_kiki').innerHTML="+chr(34)+Mochidasi_list+chr(34))
            driver.execute_script("document.getElementById('mochidashi_list').value="+chr(34)+koumoku_list[3]+chr(34))

            ##撮影物
            driver.execute_script("document.getElementById('satuei_butu_box').value="+chr(34)+koumoku_list[4]+chr(34))

            ##撮影場所
            driver.execute_script("document.getElementById('satuei_basho_box').value="+chr(34)+koumoku_list[5]+chr(34))

            ##社内
            if koumoku_list[6] == "True" :
                driver.execute_script("return document.getElementById('syanai').checked = true")
            else:
                driver.execute_script("return document.getElementById('syanai').checked = false")


            ##会議資料
            if koumoku_list[7] == "True" :
                driver.execute_script("return document.getElementById('kaigi').checked = true")
            else:
                driver.execute_script("return document.getElementById('kaigi').checked = false")


            ##保管用
            if koumoku_list[8] == "True" :
                driver.execute_script("return document.getElementById('hokan').checked = true")
            else:
                driver.execute_script("return document.getElementById('hokan').checked = false")


            ##データ移動
            if koumoku_list[9] == "True" :
                driver.execute_script("return document.getElementById('data_move').checked = true")
            else:
                driver.execute_script("return document.getElementById('data_move').checked = false")


            ##社外
            if koumoku_list[10] == "True" :
                driver.execute_script("return document.getElementById('shagai').checked = true")
            else:
                driver.execute_script("return document.getElementById('shagai').checked = false")


            ##客先
            if koumoku_list[11] == "True" :
                driver.execute_script("return document.getElementById('kyakusaki').checked = true")
            else:
                driver.execute_script("return document.getElementById('kyakusaki').checked = false")


            ##仕入先
            if koumoku_list[12] == "True" :
                driver.execute_script("return document.getElementById('siiresaki').checked = true")
            else:
                driver.execute_script("return document.getElementById('siiresaki').checked = false")


            ##その他0
            if koumoku_list[13] == "True" :
                driver.execute_script("return document.getElementById('sonota0').checked = true")
            else:
                driver.execute_script("return document.getElementById('sonota0').checked = false")


            ##INその他1
            driver.execute_script("return document.getElementById('in_sonota1').value = "+chr(34)+koumoku_list[14]+chr(34))

            ##その他
            if koumoku_list[15] == "True" :
                driver.execute_script("return document.getElementById('sonota').checked = true")
            else:
                driver.execute_script("return document.getElementById('sonota').checked = false")


            ##その他1
            print(koumoku_list[16])
            driver.execute_script("document.getElementById('sonota1').value="+chr(34)+koumoku_list[16]+chr(34))

            ##返却予定日
            driver.execute_script("document.getElementById('date2').value="+chr(34)+koumoku_list[17]+chr(34))

            ##返却予定時間
            driver.execute_script("document.getElementById('time2').value="+chr(34)+koumoku_list[18]+chr(34))

            ##ウィルスチェック
            if koumoku_list[19] == "True" :
                driver.execute_script("return document.getElementById('virus Check_box').checked = true")
            else:
                driver.execute_script("return document.getElementById('virus Check_box').checked = false")


            ##確認者
            name_sel = j.replace(name_sel, "Shouninsya_list1", "Shouninsya_list2")
            driver.execute_script("document.getElementById('Kakuninnsya').innerHTML="+chr(34)+Shouninsya_sel+"　<input type='text' size='10' style='font-size:8;' value='' id='name_text' onchange='eventA(this.id)'>"+chr(34))
            driver.execute_script("document.getElementById('Shouninsya_list1').value="+chr(34)+koumoku_list[20]+chr(34))



        if status_num == 1 or (status_num == 3 or status_num == 5) :
            ##返却日
            driver.execute_script("document.getElementById('henkyaku_bi').innerText="+chr(34)+koumoku_list[21]+chr(34))

            ##返却時間
            driver.execute_script("document.getElementById('henkyaku_jikan').innerText="+chr(34)+koumoku_list[22]+chr(34))

            ##ウィルスチェック
            if koumoku_list[23] == "True" :
                driver.execute_script("return document.getElementById('virus Check_box2').checked = true")
            else:
                driver.execute_script("return document.getElementById('virus Check_box2').checked = false")

            driver.execute_script("return document.getElementById('virus Check_box2').disabled = true")

            ##データ削除
            if koumoku_list[24] == "True" :
                driver.execute_script("return document.getElementById('Data Check_box').checked = true")
            else:
                driver.execute_script("return document.getElementById('Data Check_box').checked = false")

            driver.execute_script("return document.getElementById('Data Check_box').disabled = true")

            ##確認者2
            if status_num == 1 :
                driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML="+chr(34)+""+chr(34))
            elif status_num == 3 :
                driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML="+chr(34)+"<font size='5' color='red'>&#12958;</font>　"+name_table[koumoku_list[25]]+chr(34))


        elif status_num == 2 or status_num == 6 :

            ##返却日
            driver.execute_script("document.getElementById('date3').value="+chr(34)+today_YY4+"-"+today_mm2+"-"+today_dd2+chr(34))

            ##返却時間
            driver.execute_script("document.getElementById('time3').value="+chr(34)+today_hh2+":"+today_nn2+chr(34))

            ##確認者2
            name_sel = j.replace(Shouninsya_sel, "Shouninsya_list1", "Shouninsya_list2")
            driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML="+chr(34)+name_sel+"　<input type='text' size='10' style='font-size:8;' value='' id='name_text2' onchange='eventA(this.id)'>"+chr(34))
            if status_num == 2 : driver.execute_script("document.getElementById('Shouninsya_list2').value="+chr(34)+koumoku_list[20]+chr(34))
            if status_num == 6 : driver.execute_script("document.getElementById('Shouninsya_list2').value="+chr(34)+koumoku_list[25]+chr(34))





        ##disabled
        if status_num == 1 :
            driver.execute_script("return document.getElementById('Soushin1').disabled=true")
            driver.execute_script("return document.getElementById('Soushin2').disabled=true")
            ##driver.execute_script("return document.getElementById('Shounin1').disabled=true")
            driver.execute_script("return document.getElementById('Shounin2').disabled=true")
            ##driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")

            ##driver.execute_script("return document.getElementById('date3').disabled=true")
            ##driver.execute_script("return document.getElementById('time3').disabled=true")
            driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
            driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
            driver.execute_script("return document.getElementById('Chuushi').disabled=true")

            driver.execute_script("return document.getElementById('Shounin1').focus()")

        elif status_num == 2 :
            driver.execute_script("return document.getElementById('Soushin1').disabled=true")
            ##driver.execute_script("return document.getElementById('Soushin2').disabled=true")
            driver.execute_script("return document.getElementById('Shounin1').disabled=true")
            driver.execute_script("return document.getElementById('Shounin2').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")

            ##driver.execute_script("return document.getElementById('date3').disabled=true")
            ##driver.execute_script("return document.getElementById('time3').disabled=true")
            ##driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
            ##driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
            driver.execute_script("return document.getElementById('Chuushi').disabled=true")

            driver.execute_script("return document.getElementById('Soushin2').focus()")

        elif status_num == 3 :
            driver.execute_script("return document.getElementById('Soushin1').disabled=true")
            driver.execute_script("return document.getElementById('Soushin2').disabled=true")
            driver.execute_script("return document.getElementById('Shounin1').disabled=true")
            ##driver.execute_script("return document.getElementById('Shounin2').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")
            ##driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")

            ##driver.execute_script("return document.getElementById('date3').disabled=true")
            ##driver.execute_script("return document.getElementById('time3').disabled=true")
            driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
            driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
            driver.execute_script("return document.getElementById('Chuushi').disabled=true")

            driver.execute_script("return document.getElementById('Shounin2').focus()")

        elif status_num == 5 :
            ##driver.execute_script("return document.getElementById('Soushin1').disabled=true")
            driver.execute_script("return document.getElementById('Soushin2').disabled=true")
            driver.execute_script("return document.getElementById('Shounin1').disabled=true")
            driver.execute_script("return document.getElementById('Shounin2').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")

            ##driver.execute_script("return document.getElementById('date3').disabled=true")
            ##driver.execute_script("return document.getElementById('time3').disabled=true")
            driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
            driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
            ##driver.execute_script("return document.getElementById('Chuushi').disabled=true")
        elif status_num == 6 :
            driver.execute_script("return document.getElementById('Soushin1').disabled=true")
            ##driver.execute_script("return document.getElementById('Soushin2').disabled=true")
            driver.execute_script("return document.getElementById('Shounin1').disabled=true")
            driver.execute_script("return document.getElementById('Shounin2').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")
            driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")

            ##driver.execute_script("return document.getElementById('date3').disabled=true")
            ##driver.execute_script("return document.getElementById('time3').disabled=true")
            ##driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
            ##driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
            driver.execute_script("return document.getElementById('Chuushi').disabled=true")


        id = j.getid("情報機器持出返却管理表 - Google Chrome", "Chrome_WidgetWin_1")



        driver.execute_script("document.getElementById('Fileid').innerText="+chr(34)+"【"+file_id+"】"+chr(34))


        ##Try

        while True:

            getdata = driver.execute_script("return document.getElementById('event').innerText")

            if getdata != "" :
                driver.execute_script("document.getElementById('event').innerText="+chr(34)+""+chr(34))

                ###ラジオボタン処理
                if getdata == "syanai" :
                    driver.execute_script("return document.getElementById('kaigi').disabled = false")
                    driver.execute_script("return document.getElementById('hokan').disabled = false")
                    driver.execute_script("return document.getElementById('data_move').disabled = false")

                    driver.execute_script("return document.getElementById('kyakusaki').checked = false")
                    driver.execute_script("return document.getElementById('kyakusaki').disabled = true")
                    driver.execute_script("return document.getElementById('siiresaki').checked = false")
                    driver.execute_script("return document.getElementById('siiresaki').disabled = true")
                    driver.execute_script("return document.getElementById('sonota0').checked = false")
                    driver.execute_script("return document.getElementById('sonota0').disabled = true")
                    driver.execute_script("document.getElementById('in_sonota1').value ="+chr(34)+""+chr(34))
                    driver.execute_script("return document.getElementById('in_sonota1').disabled = true")

                    driver.execute_script("document.getElementById('sonota1').value ="+chr(34)+""+chr(34))
                    driver.execute_script("return document.getElementById('sonota1').disabled = true")

                elif getdata == "shagai" :
                    driver.execute_script("return document.getElementById('kaigi').checked=false")
                    driver.execute_script("return document.getElementById('kaigi').disabled = true")
                    driver.execute_script("return document.getElementById('hokan').checked=false")
                    driver.execute_script("return document.getElementById('hokan').disabled = true")
                    driver.execute_script("return document.getElementById('data_move').checked=false")
                    driver.execute_script("return document.getElementById('data_move').disabled = true")

                    driver.execute_script("return document.getElementById('kyakusaki').disabled = false")
                    driver.execute_script("return document.getElementById('siiresaki').disabled = false")
                    driver.execute_script("return document.getElementById('sonota0').disabled = false")
                    driver.execute_script("return document.getElementById('in_sonota1').disabled = false")

                    driver.execute_script("document.getElementById('sonota1').value ="+chr(34)+""+chr(34))
                    driver.execute_script("return document.getElementById('sonota1').disabled = true")

                elif getdata == "sonota" :
                    driver.execute_script("return document.getElementById('kaigi').checked=false")
                    driver.execute_script("return document.getElementById('kaigi').disabled = true")
                    driver.execute_script("return document.getElementById('hokan').checked=false")
                    driver.execute_script("return document.getElementById('hokan').disabled = true")
                    driver.execute_script("return document.getElementById('data_move').checked=false")
                    driver.execute_script("return document.getElementById('data_move').disabled = true")

                    driver.execute_script("return document.getElementById('kyakusaki').checked = false")
                    driver.execute_script("return document.getElementById('kyakusaki').disabled = true")
                    driver.execute_script("return document.getElementById('siiresaki').checked = false")
                    driver.execute_script("return document.getElementById('siiresaki').disabled = true")
                    driver.execute_script("return document.getElementById('sonota0').checked = false")
                    driver.execute_script("return document.getElementById('sonota0').disabled = true")
                    driver.execute_script("document.getElementById('in_sonota1').value ="+chr(34)+""+chr(34))
                    driver.execute_script("return document.getElementById('in_sonota1').disabled = true")

                    driver.execute_script("return document.getElementById('sonota1').disabled = false")



                ##確認者　入力変化
                if getdata == "name_text" :

                    getdata = driver.execute_script("return document.getElementById('name_text').value")
                    if getdata == "" :
                        continue

                    name_table_AA = {}
                    t = 0
                    for i  in range(0,len(Shouninsya_table)):
                        if j.pos(getdata, j.hash_data_out(Shouninsya_table, "hash_val", i)) > 0 :
                            name_table_AA[str(t)] = j.hash_data_out(Shouninsya_table, "hash_val", i)
                            t = t + 1

                    name_table_A=list(range(len(name_table_AA)))
                    for i in range(0,len(name_table_AA)):
                        name_table_A[i]=j.hash_data_out(name_table_AA,"hash_val",i)

                    if t == 0 :
                        pass
                    elif t == 1 :
                        getdata = name_table_A[0]
                        getdata = Shouninsya_table2[getdata]
                        driver.execute_script("return document.getElementById('Shouninsya_list1').value="+chr(34)+getdata+chr(34))
                    else:
                        getdata = j.popupmenu2(name_table_A)

                        if getdata == -1 :
                            pass
                        else:
                            getdata = name_table_A[getdata]
                            getdata = Shouninsya_table2[getdata]
                            driver.execute_script("return document.getElementById('Shouninsya_list1').value="+chr(34)+getdata+chr(34))


                ##確認者　入力変化２
                if getdata == "name_text2" :

                    getdata = driver.execute_script("return document.getElementById('name_text2').value")
                    if getdata == "" :
                        continue

                    name_table_AA = {}
                    t = 0
                    for i  in range(0,len(Shouninsya_table)):
                        if j.pos(getdata, j.hash_data_out(Shouninsya_table, "hash_val", i)) > 0 :
                            name_table_AA[str(t)] = j.hash_data_out(Shouninsya_table, "hash_val", i)
                            t = t + 1

                    name_table_A=list(range(len(name_table_AA)))
                    for i in range(0,len(name_table_AA)):
                        name_table_A[i]=j.hash_data_out(name_table_AA,"hash_val",i)


                    if t == 0 :
                        pass
                    elif t == 1 :
                        getdata = name_table_A[0]
                        getdata = Shouninsya_table2[getdata]
                        driver.execute_script("return document.getElementById('Shouninsya_list2').value="+chr(34)+getdata+chr(34))
                    else:
                        getdata = j.popupmenu2(name_table_A)

                        if getdata == -1 :
                            pass
                        else:
                            getdata = name_table_A[getdata]
                            getdata = Shouninsya_table2[getdata]
                            driver.execute_script("return document.getElementById('Shouninsya_list2').value="+chr(34)+getdata+chr(34))

                ##終了
                if getdata == "syuuryou" :
                    ##driver.Quit
                    ##Set driver = Nothing
                    Jyushin_syori = ""
                    return "shuuryou"

                if status_num == 1 : ##1回目の承認時

                    if getdata == "Shounin1" : ##承認時

                        status_num = 2

                        p1.File_write(status_num, file_id, koumoku_list[2], loginUser, koumoku_list, do_year, loginUser)

##                        件名 = "【持出前　承認】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[2]
##                        送信先CC = koumoku_list[20]
##
##                        文書 = r"情報機器持出申請（持出前）が承認されました。\n\n"
##                        文書 = 文書+"No"+file_id+r"\n"
##                        文書 = 文書+r"以下より確認をお願いします｡ \n\n"
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send

                        ##driver.Quit
                        ##Set driver = Nothing
                        ##IE_msg "承認メールを送信しました。"

                        Jyushin_syori = ""

                        ##Exit Function
                        ##GoTo_continue0=1
                        return Jyushin_syori

                    elif getdata == "Sashimodoshi1" : ##差戻時

                        ##Komento = IE_input("差戻のコメントを入力してください")
                        ##if Komento == "False" : GoTo_continue1=1

                        status_num = 5
                        p1.File_write(status_num, file_id, koumoku_list[2], loginUser, koumoku_list, do_year, loginUser)



##                        件名 = "【差戻】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[2]
##                        送信先CC = koumoku_list[20]
##
##                        文書 = r"情報機器持出申請が差し戻しされました。\n\n"
##                        文書 = 文書+"No"+file_id+r"\n"
##                        文書 = 文書+"・理由："+Komento+r"\n"
##                        文書 = 文書+r"以下より確認をお願いします｡ \n\n"
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send

                        ##driver.quit()
                        ##IE_msg "差戻メールを送信しました。"

                        Jyushin_syori = koumoku_list[2]

                        ##Exit Function
                        ###GoTo continue0
                        return Jyushin_syori




                elif status_num == 2 : ##返却時 'status

                    if getdata == "Soushin2" : ##2　送信時

                        ##読み取り処理
                        koumoku_list[21] = driver.execute_script("return document.getElementById('date3').value")##返却日
                        if koumoku_list[21] == "" :
                            j.msg("返却日を入力してください")
                            ##GoTo continue1
                            ##GoTo_continue1=1
                            ##break
                            continue

                        koumoku_list[22] = driver.execute_script("return document.getElementById('time3').value")##返却時間
                        if koumoku_list[22] == "" :
                            j.msg("返却時間を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        koumoku_list[23] = driver.execute_script("return document.getElementById('virus Check_box2').checked")       ##ウィルスチェック2
                        if koumoku_list[23] == False :
                            j.msg("ウィルスチェックをしてください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[24] = driver.execute_script("return document.getElementById('Data Check_box').checked")##データ削除
                        if koumoku_list[24] == False :
                            j.msg("データ削除してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        koumoku_list[25] = driver.execute_script("return document.getElementById('Shouninsya_list2').value")               ##確認者2
                        if koumoku_list[25] == "" :
                            j.msg("確認者を選択してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        status_num = 3

                        p1.File_write(status_num, file_id, loginUser, koumoku_list[25], koumoku_list, do_year, loginUser)

##                        件名 = "【返却】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[25]
##                        送信先CC = koumoku_list[2]
##
##                        文書 = r"情報機器持出申請の返却処理が実施されました。\n\n"
##                        文書 = 文書+"No"+file_id+r"\n"
##                        文書 = 文書+r"以下より確認をお願いします｡ \n\n"
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send
                        ##driver.quit()
                        ##Set driver = Nothing
                        ##IE_msg "返却メールを送信しました。"
                        Jyushin_syori = koumoku_list[25]
                        ##GoTo continue0
                        return Jyushin_syori
                        ##Exit Function

                elif status_num == 3 : ##返却後のサイン時 #status

                    if getdata == "Shounin2" : ##承認時

                        status_num = 4

                        p1.File_write(status_num, file_id, koumoku_list[2], loginUser, koumoku_list, do_year, loginUser)

##                        件名 = "【返却　承認】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[2]
##                        送信先CC = koumoku_list[25]
##
##                        文書 = r"情報機器持出申請が完了しました。\n\n"
##                        文書 = 文書+"No"+file_id+r"\n"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send
                        ##driver.quit()

                        ##IE_msg "返却時、承認メールを送信しました。"

                        Jyushin_syori = ""
                        ##GoTo continue0
                        return Jyushin_syori
                        ##Exit Function

                    elif getdata == "Sashimodoshi2" : ##差戻時

                        ##Komento = IE_input("差戻のコメントを入力してください")
                        ##if Komento == "False" : GoTo_continue1=1

                        status_num = 6
                        p1.File_write(status_num, file_id, koumoku_list[2], loginUser, koumoku_list, do_year, loginUser)

##                        件名 = "【差戻】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[2]
##                        送信先CC = koumoku_list[25]
##
##                        文書 = "情報機器持出申請（返却時）が差し戻しされました。"
##                        文書 = 文書+"No"+file_id+r"\n"
##                        文書 = 文書+"・理由："+Komento+r"\n"
##                        文書 = 文書+"以下より確認をお願いします｡ "
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send
                        ##driver.quit()
                        ##IE_msg "差戻メールを送信しました。"

                        Jyushin_syori = koumoku_list[2]

                        ##GoTo continue0
                        return Jyushin_syori
                        ##Exit Function




                elif status_num == 5 : ##差戻1時 'status

                    if getdata == "Soushin1" : ##2　送信時

                        ##読み取り処理
                        koumoku_list[0] = driver.execute_script("return document.getElementById('date1').value")
                        if koumoku_list[0]== "" :
                            j.msg("持出日を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        koumoku_list[1] = driver.execute_script("return document.getElementById('time1').value")
                        if koumoku_list[1] == "" :
                            j.msg("持出時刻を入力してください")
                            ##oTo_continue1=1
                            ##break
                            continue


                        koumoku_list[2] = driver.execute_script("return document.getElementById('name_list1').value")
                        if koumoku_list[2] == "" :
                            j.msg("持出者を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[3] = driver.execute_script("return document.getElementById('mochidashi_list').value")
                        if koumoku_list[3] == "" :
                            j.msg("持出機器を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[4] = driver.execute_script("return document.getElementById('satuei_butu_box').value")
##                        if koumoku_list[4] == "" and j.pos("カメ", koumoku_list[3]) > 0 :
##                            j.msg("撮影物を入力してください")
##                            ##GoTo_continue1=1
##                            ##break
##                            continue


                        koumoku_list[5] = driver.execute_script("return document.getElementById('satuei_basho_box').value")
##                        if koumoku_list[4] == "" and j.pos("カメ", koumoku_list[3]) > 0 :
##                            j.msg("撮影場所を入力してください")
##                            ##GoTo_continue1=1
##                            ##break
##                            continue



                        koumoku_list[6] = driver.execute_script("return document.getElementById('syanai').checked")
                        if koumoku_list[6] == True :
                            koumoku_list[7] = driver.execute_script("return document.getElementById('kaigi').checked")
                            koumoku_list[8] = driver.execute_script("return document.getElementById('hokan').checked")
                            koumoku_list[9] = driver.execute_script("return document.getElementById('data_move').checked")

                            if koumoku_list[7] == False and (koumoku_list[8] == False and koumoku_list[9] == False) :
                                j.msg("社内のチェックボックスを入力してください")
                                ##GoTo_continue1=1
                                ##break
                                continue

                        else:
                            koumoku_list[10] = driver.execute_script("return document.getElementById('shagai').checked")
                            if koumoku_list[10] == True :
                                koumoku_list[11] = driver.execute_script("return document.getElementById('kyakusaki').checked")
                                koumoku_list[12] = driver.execute_script("return document.getElementById('siiresaki').checked")
                                koumoku_list[13] = driver.execute_script("return document.getElementById('sonota0').checked")

                                if koumoku_list[11] == False and (koumoku_list[12] == False and koumoku_list[13] == False) :
                                    j.msg("社外のチェックボックスを入力してください")
                                    ##GoTo_continue1=1
                                    ##break
                                    continue


                                if koumoku_list[13] == True :
                                    koumoku_list[14] = driver.execute_script("return document.getElementById('in_sonota1').value")
                                    if koumoku_list[14] == "" :
                                        j.msg("その他を入力してください")
                                        ##GoTo_continue1=1
                                        ##break
                                        continue



                            else:
                                koumoku_list[15] = driver.execute_script("return document.getElementById('sonota').checked")
                                if koumoku_list[15] == True :
                                    koumoku_list[16] = driver.execute_script("return document.getElementById('sonota1').value")
                                    if koumoku_list[16] == "" :
                                        j.msg("その他を入力してください")
                                        ##GoTo_continue1=1
                                        ##break
                                        continue

                                else:
                                    j.msg("使用目的を入力してください")
                                    ##GoTo_continue1=1
                                    ##break
                                    continue




                        koumoku_list[17] = driver.execute_script("return document.getElementById('date2').value")
                        if koumoku_list[17] == "" :
                            j.msg("返却予定日を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[18] = driver.execute_script("return document.getElementById('time2').value")
                        if koumoku_list[18] == "" :
                            j.msg("返却予定時間を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[19] = driver.execute_script("return document.getElementById('virus Check_box').checked")
                        if koumoku_list[19] == False :
                            j.msg("ウィルスチェックをしてください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[20] = driver.execute_script("return document.getElementById('Shouninsya_list1').value")
                        if koumoku_list[20] == "" :
                            j.msg("確認者を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        if koumoku_list[20] == koumoku_list[2] :
                            j.msg("持出者と確認者が同じです。")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[21] = ""
                        koumoku_list[22] = ""
                        koumoku_list[23] = ""
                        koumoku_list[24] = ""
                        koumoku_list[25] = ""


                        status_num = 1

                        p1.File_write(status_num, file_id, loginUser, koumoku_list[20], koumoku_list, do_year, loginUser)

##                        件名 = "【申請】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[20]
##                        送信先CC = koumoku_list[2]
##
##                        文書 = "情報機器持出申請です。"
##                        文書 = 文書+"№"+file_id+r"\n"
##                        文書 = 文書+"以下より確認をお願いします｡ "
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send
                        ##IE_msg "申請メールを送信しました。"
                        ##driver.quit()

                        Jyushin_syori = koumoku_list[20]
                        ##GoTo continue0
                        return Jyushin_syori
                        ##Exit Function

                    elif getdata == "Chuushi" :             ##中止
                        status_num = 7

                        p1.File_write(status_num, file_id, loginUser, koumoku_list[20], koumoku_list, do_year, loginUser)
                        j.msg("本件は中止しました。")

                        ##driver.quit()
                        return ""




                elif status_num == 6 : ##差戻2時 'status

                    if getdata == "Soushin2" : ##2　送信時

                        #読み取り処理
                        koumoku_list[21] = driver.execute_script("return document.getElementById('date3').value")                    ##返却日
                        if koumoku_list[21] == "" :
                            j.msg("返却日を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        koumoku_list[22] = driver.execute_script("return document.getElementById('time3').value")                    ##返却時間
                        if koumoku_list[22] == "" :
                            j.msg("返却時間を入力してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[23] = driver.execute_script("return document.getElementById('virus Check_box2').checked")       ##ウィルスチェック2
                        if koumoku_list[23] == False :
                            j.msg("ウィルスチェックをしてください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[24] = driver.execute_script("return document.getElementById('Data Check_box').checked")         ##データ削除
                        if koumoku_list[24] == False :
                            j.msg("データ削除してください")
                            ##GoTo_continue1=1
                            ##break
                            continue


                        koumoku_list[25] = driver.execute_script("return document.getElementById('Shouninsya_list2').value")               ##確認者2
                        if koumoku_list[25] == "" :
                            j.msg("確認者を選択してください")
                            ##GoTo_continue1=1
                            ##break
                            continue

                        if koumoku_list[25]==koumoku_list[2]:
                            j.msg("持出者と確認者が同じです")
                            continue



                        status_num = 3

                        p1.File_write(status_num, file_id, loginUser, koumoku_list[25], koumoku_list, do_year, loginUser)

##                        件名 = "【返却】情報機器持出申請 "+file_id
##                        送信先 = koumoku_list[25]
##                        送信先CC = koumoku_list[2]
##
##                        文書 = "情報機器持出申請の返却処理が実施されました。"
##                        文書 = 文書+"No"+file_id+r"\n"
##                        文書 = 文書+"以下より確認をお願いします｡ "
##                        文書 = 文書+"<"+exe_path+">"

                        send = 2
                        ##sendmail 件名, 送信先, 送信先CC, 文書, send
                        ##IE_msg "返却メールを送信しました。"
                        Jyushin_syori = koumoku_list[25]
                        ##driver.quit()
                        return Jyushin_syori

        if GoTo_continue0 == 1 : break
        ##continue1:
    ##finish:


    ##    driver.quit()
    ##    Set driver = Nothing



if __name__ == '__main__':
    main()
