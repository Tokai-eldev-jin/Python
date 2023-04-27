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

import jin as j
import global_value as g


def main():
    pass




def Create_Table(loginUser, kind,name_table,driver):


    import inputbox2 as p1
    import save_to_file as p2


    ##kind :"personal"⇒個人　　"all"⇒全部

    i:int
    ii:int
    t:int
    do_path:str
    page_table:str
    kazu:int
    table_str1:str
    table_str2:str
    status_num:int
    cnt:int
    Hani:str
    Hani_d:int
    Hani_u:int
    year_num:str


    continue0=0
    while True:
        if continue0==0:
            Hani = "1-32678"
            g.token_str=Hani
            Hani_d = int(j.token("-", g.token_str))
            Hani_u = int(j.token("-", g.token_str))



            if kind == "personal":
                do_path = j.path_join(g.CUR_DIR , "user" , loginUser , loginUser + ".txt")
            else:
                do_path = j.path_join(g.CUR_DIR , "ini")

                ##年のデータファイルの数を数える
                kazu = j.getdir(do_path, "*_datafile.txt")

                ##年のデータファイルを配列に入れる
                year_table=j.Hairetu(kazu)

                for i in range(0 , kazu):
                    year_table[i] = j.replace(g.getdir_files[i], "_datafile.txt", "")


                ##年をソートする
                year_table = j.Hairetu_sort(year_table)

                ##年のタグソースを作成する
                page_table = "<select id='year_sen' onchange='eventA(this.id)'>"
                for i in range(0 , kazu):
                    page_table = page_table + "<option id='" + year_table[i] + "'>" + year_table[i] + "</option>"

                page_table = page_table + "</select>"

                do_path = j.path_join(g.CUR_DIR , "ini" , year_table[0]+"_datafile.txt")
                year_num = year_table[0]

        ##continue0:
        continue0=0



        if kind == "personal":
            table_str1 = "<table id='table1'><caption>" + name_table[loginUser] + "</caption>"
        else:
            table_str1 = "<table id='table1'><caption>" + "ALL" + "</caption>"


        table_str1 = table_str1 + "<th width='40'>番号</th>"
        table_str1 = table_str1 + "<th width='80'>状態</th>"
        table_str1 = table_str1 + "<th width='90'>持出日</th>"
        table_str1 = table_str1 + "<th width='50'>持出時刻</th>"
        table_str1 = table_str1 + "<th width='100'>持出者</th>"
        table_str1 = table_str1 + "<th width='200'>持出機器</th>"
        table_str1 = table_str1 + "<th width='150'>撮影物</th>"
        table_str1 = table_str1 + "<th width='100'>撮影場所</th>"
        table_str1 = table_str1 + "<th width='250'>使用目的</th>"
        table_str1 = table_str1 + "<th width='90'>返却予定日</th>"
        table_str1 = table_str1 + "<th width='50'>返却予定時間</th>"
        table_str1 = table_str1 + "<th width='40'>virus Check</th>"
        table_str1 = table_str1 + "<th width='100'>確認者</th>"
        table_str1 = table_str1 + "<th width='90'>返却日</th>"
        table_str1 = table_str1 + "<th width='50'>返却時間</th>"
        table_str1 = table_str1 + "<th width='40'>virus Check</th>"
        table_str1 = table_str1 + "<th width='40'>データ削除</th>"
        table_str1 = table_str1 + "<th width='100'>確認者</th>"


        ##ファイルに案件がなければ終了する
        str_table = {}
        f1 = j.fopen(do_path, "f_read")
        kazu = j.fget(f1, -1)

        if kazu == 0:
            j.fclose(f1)

            j.msg("案件がありません。")

            break




        ##タグソースを作成する
        cnt = 0
        for i in range(1 , kazu+1):

            getdata = j.fget(f1, i, 3)                                ##statusの取得
            if getdata == "": break
            getdata=int(getdata)

            if getdata == 7:
                ##GoTo_continue2=1                   ##中止は無視する
                continue

            if int(j.fget(f1, i, 2)) < Hani_d: continue          ##番号で抽出する
            if int(j.fget(f1, i, 2)) > Hani_u: continue         ##番号で抽出する


            ##1行だけのタグソースを作成する
            table_str2 = "<tr>"

            file_id = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)         #file-id
            table_str2 = table_str2 + "<td><a href='javascript:void(0)' id='link_" + file_id + "' onclick='eventA(this.id)'>" + file_id + "</a></td>"

            status_num = int(j.fget(f1, i, 3))                             ##status
            getdata2 = j.fget(f1, i, 4 + 26)                          ##完了日を得る
            if getdata2 != "": getdata2 = "<br>" + j.fget(f1, i, 4 + 26) + "<br>" + j.fget(f1, i, 4 + 27)   ##完了日と完了時間

            if status_num == 1:
                getdata = "申請中"
                table_str2 = table_str2 + "<td bgcolor='lightgreen'>" + getdata + "</td>"
            elif status_num == 2:
                getdata = "持出"
                table_str2 = table_str2 + "<td bgcolor='yellow'>" + getdata + "</td>"
            elif status_num == 3:
                getdata = "返却"
                table_str2 = table_str2 + "<td bgcolor='orange'>" + getdata + "</td>"
            elif status_num == 4:
                getdata = "完了"
                table_str2 = table_str2 + "<td bgcolor='silver'>" + getdata + getdata2 + "</td>"
            elif status_num == 5:
                getdata = "差戻（持出）"
                table_str2 = table_str2 + "<td bgcolor='pink'>" + getdata + "</td>"
            elif status_num == 6:
                getdata = "差戻（返却）"
                table_str2 = table_str2 + "<td bgcolor='pink'>" + getdata + "</td>"



            t = 0
            for ii in range( 0 , 26):
                getdata = j.fget(f1, i, 4 + t)

                if t >= 6 and t <= 16:
                    if t == 6 and getdata == "True":

                        table_str2 = table_str2 + "<td>社内"
                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/会議資料"

                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/保管用"

                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/データ移動"


                        table_str2 = table_str2 + "</td>"

                        t = 16
                    elif t == 10 and getdata == "True":
                        table_str2 = table_str2 + "<td>社外"
                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/客先"

                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/仕入先"

                        t = t + 1

                        if j.fget(f1, i, 4 + t) == "True":
                            table_str2 = table_str2 + "/" + j.fget(f1, i, 5 + t)

                        table_str2 = table_str2 + "</td>"

                        t = 16
                    elif t == 15 and getdata == "True":
                        table_str2 = table_str2 + "<td>" + j.fget(f1, i, 5 + t) + "</td>"

                else:
                    ##チェックマーク
                    if t == 19 and getdata == "True": getdata = "&#x2611;"
                    if t == 19 and getdata == "False": getdata = "&#x2610;"
                    if t == 23 and getdata == "True": getdata = "&#x2611;"
                    if t == 23 and getdata == "False": getdata = "&#x2610;"
                    if t == 24 and getdata == "True": getdata = "&#x2611;"
                    if t == 24 and getdata == "False": getdata = "&#x2610;"


                    if t == 2: getdata = name_table[getdata]

                    ##印マーク
                    if t == 20:
                        if status_num == 1 or status_num == 5: getdata = name_table[getdata]
                        if status_num == 2: getdata = "<font size='2' color='red'>&#12958;</font>" + name_table[getdata]
                        if status_num == 3: getdata = "<font size='2' color='red'>&#12958;</font>" + name_table[getdata]
                        if status_num == 4: getdata = "<font size='2' color='red'>&#12958;</font>" + name_table[getdata]
                        if status_num == 6: getdata = "<font size='2' color='red'>&#12958;</font>" + name_table[getdata]


                    ##印マーク
                    if t == 25:
                        if status_num == 4:
                            getdata = "<font size='2' color='red'>&#12958;</font>" + name_table[getdata]
                        else:
                            if j.hash_data_out(name_table,"hash_exists",getdata) == True:
                                getdata = name_table[getdata]
                            else:
                                getdata=""



                    table_str2 = table_str2 + "<td>" + getdata + " </td>"


                t = t + 1
                if t > 25: break



            table_str2 = table_str2 + "</tr>"
            str_table[cnt] = table_str2
            cnt = cnt + 1


        ##continue2




        if cnt == 0:
            ##id = Application.hWnd
            j.msg("案件がありません。")
            break



        ##タグソースを合成する
        table_str1 = table_str1 + str_table[0]
        for i in range(1 , len(str_table) ):
            table_str1 = table_str1 + str_table[i]

        table_str1 = table_str1 + "</table>"






        ############ Chrome 起動　###########
        driver.maximize_window()
        driver.get(j.path_join(g.CUR_DIR , "ini","Table.html"))
        j.sleep (200)
        j.BusyWait(driver)
        driver.refresh()
        ############ Chrome 起動　###########


        ##Chomeに表示
        if kind == "personal":
            pass
            ##driver.execute_script("document.getElementById('kind').innerText=" + chr(34) + name_table[loginUser] + chr(34))
        else:
            driver.execute_script("document.getElementById('year').innerHTML = " + chr(34) + page_table + chr(34))
            driver.execute_script("return document.getElementById('year').style.display='block'")
            driver.execute_script("document.getElementById('year_sen').value=" + year_num)
            ##driver.execute_script("document.getElementById('kind').innerText = " + chr(34) + "ALL" + chr(34))


        driver.execute_script("document.getElementById('table_in').innerHTML = " + chr(34) + table_str1 + chr(34))





        ##id = Application.hWnd


        while True:

            getdata = driver.execute_script("return document.getElementById('event').innerText")

            if getdata != "":
                driver.execute_script("document.getElementById('event').innerText=" + chr(34) + "" + chr(34))

                ##終了時
                if getdata == "end":
                    ##driver.Quit
                    ##Set driver = Nothing
                    return ""

                ##年選択時
                elif getdata == "year_sen":
                    getdata = driver.execute_script("return document.getElementById('year_sen').value")
                    do_path = j.path_join(g.CUR_DIR , "ini" , getdata + "_datafile.txt")
                    driver.execute_script("document.getElementById('table_in').innerHTML=" + chr(34) + "" + chr(34))
                    year_num = getdata

                    ##抽出をリセットする
                    Hani = "1-32678"
                    g.token_str=Hani
                    Hani_d = int(j.token("-", g.token_str))
                    Hani_u = int(j.token("-", g.token_str))

                    GoTo_continue0=1
                    break



                ##Link表示
                if j.copy(getdata, 1, 4) == "link":
                    file_id = j.copy(getdata, 6, len(getdata))
                    getdata=Hyouji(name_table, file_id,driver)

                    if j.pos("Sashimodoshi",getdata)>0:
                        return "Sashimodoshi"
                    else:
                        break



                ##抽出
                if getdata == "chuusyutu":
                    getdata = p1.inputbox2("範囲を入力して下さい。")
                    if getdata != False:
                        if j.pos("-", getdata) == 0:
                            j.msg(chr(34) + "-" + chr(34) + "を含めてください")
                            ##GoTo continue1
                            continue
                        elif j.IsNum(j.copy(getdata, 1, j.pos("-", getdata) - 1)) == False:
                            j.msg("数字を入力してください")
                            ##GoTo continue1
                            continue
                        elif j.IsNum(j.copy(getdata, j.pos("-", getdata) + 1, len(getdata))) == False:
                            j.msg("数字を入力してください")
                            ##GoTo continue1
                            continue

                        Hani = getdata

                        g.token_str = Hani
                        Hani_d = int(j.token("-", g.token_str))
                        Hani_u = int(j.token("-", g.token_str))

                        if Hani_d > Hani_u:
                            j.msg("前の数字が後の数字より大きいです。")
                            ##GoTo continue1
                            continue
                        continue0=1
                        break

                ##excelに保存
                if getdata == "hozon":

                    p2.save_to_file(driver)

            ##continue1:








def Hyouji(name_table,file_id,driver):
    from selenium import webdriver
    import chromedriver_autoinstaller
    import File_write as p1

##    options = webdriver.ChromeOptions()
##    options.add_experimental_option("excludeSwitches", ['enable-automation'])

    nenn:str
    kazu:int
    i:int
    t:int
    status_num:int


    nenn = j.copy(file_id, 1, 4)

    f1={}
    f1 = j.fopen(j.path_join(g.CUR_DIR , "ini" , nenn + "_datafile.txt"), "f_read")
    kazu = j.fget(f1, -1)
    koumoku_list = {}

    for i in range(1 , kazu+1):
        getdata = j.fget(f1, i, 1) + "_" + j.fget(f1, i, 2)


        if getdata == file_id:

            status_num = int(j.fget(f1, i, 3))

            for t in range( 0 , 26):
                koumoku_list[t] = j.fget(f1, i, 4 + t)


            break



    j.fclose(f1)





    ########### Chrome 起動　###########
##    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    ##driver.set_window_size(g.g_screen_w/5, g.g_screen_h/3)
    driver.get(j.path_join(g.CUR_DIR,r"ini\index.html"))
    j.BusyWait(driver)
    j.sleep(100)
    driver.refresh()
    ########### Chrome 起動　###########




    ##持出日
    driver.execute_script("document.getElementById('Mochidashi_bi').innerText=" + chr(34) + koumoku_list[0] + chr(34))

    ##持出時刻
    driver.execute_script("document.getElementById('Mochidashi_jikoku').innerText=" + chr(34) + koumoku_list[1] + chr(34))

    ##持出者
    driver.execute_script("document.getElementById('mochidashi_sya').innerText=" + chr(34) + name_table[koumoku_list[2]] + chr(34))

    ##持出機器
    driver.execute_script("document.getElementById('mochidashi_kiki').innerText=" + chr(34) + koumoku_list[3] + chr(34))

    ##撮影物
    driver.execute_script("document.getElementById('satuei_butu_box').value=" + chr(34) + koumoku_list[4] + chr(34))
    driver.execute_script("return document.getElementById('satuei_butu_box').disabled=true")

    ##撮影場所
    driver.execute_script("document.getElementById('satuei_basho_box').value=" + chr(34) + koumoku_list[5] + chr(34))
    driver.execute_script("return document.getElementById('satuei_basho_box').disabled=true")

    ##社内
    if koumoku_list[6] == "True":
        driver.execute_script("return document.getElementById('syanai').checked = true")
    else:
        driver.execute_script("return document.getElementById('syanai').checked = false")

    driver.execute_script("return document.getElementById('syanai').disabled = true")

    ##会議資料
    if koumoku_list[7] == "True":
        driver.execute_script("return document.getElementById('kaigi').checked = true")
    else:
        driver.execute_script("return document.getElementById('kaigi').checked = false")

    driver.execute_script("return document.getElementById('kaigi').disabled = true")

    ##保管用
    if koumoku_list[8] == "True":
        driver.execute_script("return document.getElementById('hokan').checked = true")
    else:
        driver.execute_script("return document.getElementById('hokan').checked = false")

    driver.execute_script("return document.getElementById('hokan').disabled = true")

    ##データ移動
    if koumoku_list[9] == "True":
        driver.execute_script("return document.getElementById('data_move').checked = true")
    else:
        driver.execute_script("return document.getElementById('data_move').checked = false")

    driver.execute_script("return document.getElementById('data_move').disabled = true")

    ##社外
    if koumoku_list[10] == "True":
        driver.execute_script("return document.getElementById('shagai').checked = true")
    else:
        driver.execute_script("return document.getElementById('shagai').checked = false")

    driver.execute_script("return document.getElementById('shagai').disabled = true")

    ##客先
    if koumoku_list[11] == "True":
        driver.execute_script("return document.getElementById('kyakusaki').checked = true")
    else:
        driver.execute_script("return document.getElementById('kyakusaki').checked = false")

    driver.execute_script("return document.getElementById('kyakusaki').disabled = true")

    ##仕入先
    if koumoku_list[12] == "True":
        driver.execute_script("return document.getElementById('siiresaki').checked = true")
    else:
        driver.execute_script("return document.getElementById('siiresaki').checked = false")

    driver.execute_script("return document.getElementById('siiresaki').disabled = true")

    ##その他0
    if koumoku_list[13] == "True":
        driver.execute_script("return document.getElementById('sonota0').checked = true")
    else:
        driver.execute_script("return document.getElementById('sonota0').checked = false")

    driver.execute_script("return document.getElementById('sonota0').disabled = true")

    ##INその他1
    driver.execute_script("return document.getElementById('in_sonota1').value = " + chr(34) + koumoku_list[14] + chr(34))
    driver.execute_script("return document.getElementById('in_sonota1').disabled = true")

    ##その他
    if koumoku_list[15] == "True":
        driver.execute_script("return document.getElementById('sonota').checked = true")
    else:
        driver.execute_script("return document.getElementById('sonota').checked = false")

    driver.execute_script("return document.getElementById('sonota').disabled = true")

    ##その他1
    driver.execute_script("document.getElementById('sonota1').value=" + chr(34) + koumoku_list[16] + chr(34))
    driver.execute_script("document.getElementById('sonota1').disabled=true")

    ##返却予定日
    driver.execute_script("document.getElementById('henkyaku_yoteibi').innerText=" + chr(34) + koumoku_list[17] + chr(34))

    ##返却予定時間
    driver.execute_script("document.getElementById('henkyaku_yoteijikan').innerText=" + chr(34) + koumoku_list[18] + chr(34))

    ##ウィルスチェック
    if koumoku_list[19] == "True":
        driver.execute_script("return document.getElementById('virus Check_box').checked = true")
    else:
        driver.execute_script("return document.getElementById('virus Check_box').checked = false")

    driver.execute_script("return document.getElementById('virus Check_box').disabled = true")

    ##確認者
    if j.hash_data_out(name_table,"hash_exists",koumoku_list[20]) != "":
        if status_num == 2 or (status_num == 3 or (status_num == 4 or status_num == 6)):
            driver.execute_script("document.getElementById('Kakuninnsya').innerHTML=" + chr(34) + "<font size='5' color='red'>&#12958;</font>　" + name_table[koumoku_list[20]] + chr(34))
        else:
            driver.execute_script("document.getElementById('Kakuninnsya').innerHTML=" + chr(34) + name_table[koumoku_list[20]] + chr(34))

    else:
        driver.execute_script("document.getElementById('Kakuninnsya').innerHTML=" + chr(34) + "" + chr(34))


    ##返却日
    driver.execute_script("document.getElementById('henkyaku_bi').innerText=" + chr(34) + koumoku_list[21] + chr(34))

    ##返却時間
    driver.execute_script("document.getElementById('henkyaku_jikan').innerText=" + chr(34) + koumoku_list[22] + chr(34))

    ##ウィルスチェック
    if koumoku_list[23] == "True":
        driver.execute_script("return document.getElementById('virus Check_box2').checked = true")
    else:
        driver.execute_script("return document.getElementById('virus Check_box2').checked = false")

    driver.execute_script("return document.getElementById('virus Check_box2').disabled = true")

    ##データ削除
    if koumoku_list[24] == "True":
        driver.execute_script("return document.getElementById('Data Check_box').checked = true")
    else:
        driver.execute_script("return document.getElementById('Data Check_box').checked = false")

    driver.execute_script("return document.getElementById('Data Check_box').disabled = true")

    ##確認者2
    if j.hash_data_out(name_table,"hash_exists",koumoku_list[25]) == False:
        driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML=" + chr(34) + "" + chr(34))
    else:
        if status_num == 4:
            driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML=" + chr(34) + "<font size='5' color='red'>&#12958;</font>　" + name_table[koumoku_list[25]] + chr(34))
        else:
            driver.execute_script("document.getElementById('Kakuninnsya2').innerHTML=" + chr(34) + name_table[koumoku_list[25]] + chr(34))




    driver.execute_script("return document.getElementById('Soushin1').disabled=true")
    driver.execute_script("return document.getElementById('Soushin2').disabled=true")
    driver.execute_script("return document.getElementById('Shounin1').disabled=true")
    driver.execute_script("return document.getElementById('Shounin2').disabled=true")


    driver.execute_script("return document.getElementById('virus Check_box2').disabled=true")
    driver.execute_script("return document.getElementById('Data Check_box').disabled=true")
    driver.execute_script("return document.getElementById('Chuushi').disabled=true")

    if status_num == 1:
        driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=false")
        driver.execute_script("document.getElementById('Sashimodoshi1').value=" + chr(34) + "自分で差戻" + chr(34))
    else:
        driver.execute_script("return document.getElementById('Sashimodoshi1').disabled=true")


    if status_num == 3:
        driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=false")
        driver.execute_script("document.getElementById('Sashimodoshi2').value=" + chr(34) + "自分で差戻" + chr(34))
    else:
        driver.execute_script("return document.getElementById('Sashimodoshi2').disabled=true")


    driver.execute_script("document.getElementById('Fileid').innerText=" + chr(34) + "【" + file_id + "】" + chr(34))



    ##try:

    while True:

        getdata = driver.execute_script("return document.getElementById('event').innerText")

        if getdata != "":
            driver.execute_script("document.getElementById('event').innerText=" + chr(34) + "" + chr(34))


            if getdata == "syuuryou":

##                driver.quit()
                return ""

            elif getdata == "Sashimodoshi1":

                持出者 = koumoku_list[2]
                確認者 = koumoku_list[20]
                status_num = 5

                p1.File_write(status_num, file_id, 持出者, 確認者, koumoku_list, nenn, 持出者)
##                driver.quit()
                return "Sashimodoshi1"


            elif getdata == "Sashimodoshi2":
                持出者 = koumoku_list[2]
                確認者 = koumoku_list[25]
                status_num = 6

                p1.File_write(status_num, file_id, 持出者, 確認者, koumoku_list, nenn, 持出者)

##                driver.quit()
                return "Sashimodoshi2"

##    driver.quit()






if __name__ == '__main__':
    main()
