def main():

    import jin as j
    import global_value as g

    import os
    import sys
    import time

    #web
    #必要なライブラリをインポート
    from selenium import webdriver
    import chromedriver_autoinstaller
    chromedriver_autoinstaller.install()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])

    j.get_gamen_size()

    def get_CUR_DIR():
        import sys
        import os

        path_current_dir = os.path.dirname(sys.argv[0])

        #print(sys.argv[0])
        #print(path_current_dir)
        return  path_current_dir

    g.CUR_DIR=get_CUR_DIR()

    #print(sys.argv[0])#起動パラメータ

    #グローバル変数にはg.を付ける







    #importファイル
    import name_list_create as p1
    import new_create as p2
    import Jyushin_syori as p4
    import Create_Table as p5





    GoTo_Start=0
    GoTo_continue1=0

    driver = webdriver.Chrome(options=options)

    while True:


        i:int
        loginUser:str
        password:str
        in_password:str
        kazu:int
        Shiyou_mokuteki:str
        Menu_list:str
        Jyushin_flag:int
        DB_path1:str
        name_sel:str
        id:int
        next_login:str



        if (GoTo_Start==0 and GoTo_continue1==0):
            next_login = ""
            start_flag=1


            ##id = Application.hWnd

            ##'##########　従業員番号と名前のtableを作成する　##########
            DB_path1 = j.path_join(g.CUR_DIR,r"ini\NAME.csv")
            name_table = {}
            name_table2 = {}

            name_sel = "<select id='name_list1'>"

            name_table1 = p1.name_list_create(DB_path1)

            for i in range(0,len(name_table1)):
                name_table[name_table1[i][1]] = name_table1[i][2]   #従業員番号=名前
                name_table2[name_table1[i][2]] = name_table1[i][1]  #名前=従業員番号
                name_sel = name_sel+"<option value='"+name_table1[i][1]+"'>"+name_table1[i][2]+"</option>" #value=従業員番号
            name_sel = name_sel+"</select>"

            ##print(name_sel)
            ###########　従業員番号と名前のtableを作成する　##########


            #########　承認者のtableを作成する　##########
            Shouninsya_path:str
            Shouninsya_sel:str

            Shouninsya_path = j.path_join(g.CUR_DIR,r"ini\Shounin_list.txt")
            Shouninsya_table = {}
            Shouninsya_table2 = {}

            Shouninsya_sel = "<select id='Shouninsya_list1'>"
            f1={}
            f1 = j.fopen(Shouninsya_path, "f_read")
            kazu = j.fget(f1, -1)
            for i in range(1,kazu+1):
                Shouninsya_table[j.fget(f1, i, 1)] = j.fget(f1, i, 2) #従業員番号=名前
                Shouninsya_table2[j.fget(f1, i, 2)] = j.fget(f1, i, 1) #名前=従業員番号
                Shouninsya_sel = Shouninsya_sel+"<option value='"+j.fget(f1, i, 1)+"'>"+j.fget(f1, i, 2)+"</option>" #value=従業員番号

            Shouninsya_sel = Shouninsya_sel+"</select>"
            j.fclose(f1)
            ##########　承認者のtableを作成する　##########

        ##    Start:
        GoTo_Start=0

        if GoTo_continue1==0:
            start_flag=1

            ########### chrome　立ち上げ ##########
            ##    driver.maximize_window()
            driver.set_window_size(g.g_screen_w/5, g.g_screen_h/3)
            driver.get(j.path_join(g.CUR_DIR,r"ini\login.html"))
            j.BusyWait(driver)
            j.sleep(100)
            driver.refresh()

            j.sleep (200)
            if next_login == "":
                driver.execute_script("document.getElementById('user').focus()")
            else:
                driver.execute_script("document.getElementById('user').value="+chr(34)+next_login+chr(34))
                driver.execute_script("return document.getElementById('login').focus()")


            ########### ログイン画面
            while True:

                getdata = driver.execute_script("return document.getElementById('event').innerText")

                if getdata != "":
                    driver.execute_script("document.getElementById('event').innerText="+chr(34)+""+chr(34))

                    if getdata == "login":
                        loginUser = driver.execute_script("return document.getElementById('user').value")

                        #従業員番号が正しくないとき
                        if j.hash_data_out(name_table,"hash_exists",loginUser) == False:
                            j.msg("正しい従業員番号を入力してください。", "情報機器持出管理",0)

                        #従業員番号が正しいとき
                        else:
                            #承認者の場合
                            if j.hash_data_out(Shouninsya_table,"hash_exists",loginUser) ==True :

                                f1 = j.fopen(j.path_join(g.CUR_DIR,r"ini\Shounin_list.txt"), "f_read")
                                kazu = j.fget(f1, -1)
                                for i in range(1,kazu+1):
                                    getdata = j.fget(f1, i, 1)
                                    if getdata == loginUser :
                                        password = j.fget(f1, i, 3)
                                        break

                                j.fclose(f1)


                                #Passwordが登録されていないとき
                                if password == "" :
                                    #Passwordの新規登録
                                    password = j.IE_input("Passwordの登録をお願いします","", False, True)

                                    if password != "" and password != "False" :
                                        f1 = j.fopen(j.path_join(g.CUR_DIR,r"ini\Shounin_list.txt"), "f_read_or_f_write")
                                        kazu = j.fget(f1, -1)
                                        for i in range(1,kazu+1):
                                            getdata = j.fget(f1, i, 1)
                                            if getdata == loginUser :
                                                j.fput(f1, password, i, 3)
                                                break
                                        j.fclose(f1)

                                        break

                                    else:
                                        j.sleep(10)

                                #Passwordが登録されている
                                else:
                                    #Password入力欄を出す
                                    driver.execute_script("document.getElementById('div2').style.visibility='visible'")
                                    driver.execute_script ("return document.getElementById('password').focus()")
                                    in_password = driver.execute_script("return document.getElementById('password').value")

                                    if in_password != "" :
                                        #Passwordが一致した場合
                                        if in_password == password :
                                            break
                                        else:
                                            j.msg("Passwordが違います。")
                                    else:
                                        j.sleep(10)

                            #承認者ではない場合
                            else:
                                break

                    elif getdata == "end" :
                        ##'baloon id, "終了しました。", "excel"
                        driver.quit()
                        j.End()
                j.sleep (10)






            #    loginUser = get_user()
            #    loginUser = vb_strconv(loginUser, 1)
            #    loginUser = "A2033"
            #    loginUser = "A4134"

            ##########　ユーザーフォルダを作成する　##########
            if j.fopen(j.path_join(g.CUR_DIR,"user",loginUser), "f_exists") == False :
                j.File_system("fs_createfolder", j.path_join(g.CUR_DIR,"user",loginUser))
            ##########　ユーザーフォルダを作成する　##########

            Jyushin_flag = 0

        ##continue1:
        GoTo_continue1=0


        Menu_list = name_table[loginUser]+"さん。メニューを選択してください。<br><br>"
        Menu_list = Menu_list + "<select id='menu' onchange='eventA(this.value)'><option id='blank' value=''></option>"



        ##########　受信の確認　##########

        if j.fopen(j.path_join(g.CUR_DIR,"user",loginUser,"Re_"+loginUser+".txt"), "f_exists") == True :

            f1 = j.fopen(j.path_join(g.CUR_DIR,"user",loginUser,"Re_"+loginUser+".txt"), "f_read")
            kazu = j.fget(f1, -1)

            for i in range(1,kazu+1):
                getdata = int(j.fget(f1, i, 3))
                if getdata != 0 and getdata != 4 :
                    Menu_list = Menu_list+"<option id='jyusin' value='jyusin'>受信BOX</option>"
                    if Jyushin_flag == 0 : Jyushin_flag = 1
                    break
            j.fclose(f1)
        else:
            pass







        ##########　受信フォルダにファイルがある場合　##########

        Menu_list = Menu_list+"<option id='shinki' value='shinki'>新規作成</option>"
        Menu_list = Menu_list+"<option id='shinki' value='jissi'>自分の案件</option>"
        Menu_list = Menu_list+"<option id='hyou' value='hyou'>すべての案件</option>"
        Menu_list = Menu_list+"<option id='modoru' value='modoru'>ログインに戻る</option></select>"

        #受信がある場合は受信を開く
        if Jyushin_flag == 1 and start_flag==1:

            #受信あるときは受信を開く
            next_login = p4.Jyushin_syori(loginUser, name_sel, name_table, name_table2, Shouninsya_sel, Shouninsya_table, Shouninsya_table2, driver)

            if next_login == "shuuryou":
                yushin_flag = 2
            elif next_login != "":
                Jyushin_flag = 0
                GoTo_Start=1
                continue
            else:
                Jyushin_flag = 2
                GoTo_Start=1
                continue

        start_flag=0
        Jyushin_flag=0

        ########### Chrome 起動　###########
        ##    driver.maximize_window()
        driver.set_window_size(g.g_screen_w/5, g.g_screen_h/3)
        driver.get(j.path_join(g.CUR_DIR,r"ini\menu.html"))
        j.BusyWait(driver)
        j.sleep(100)
        driver.refresh()
        ########### Chrome 起動　###########

        #menuを表示する
        driver.execute_script("document.getElementById('div1').innerHTML="+chr(34)+Menu_list+chr(34))
        j.sleep (100)

        driver.execute_script("return document.getElementById('menu').focus()")


##        try:

        while True:

            getdata = driver.execute_script("return document.getElementById('event').innerText")

            if getdata != "" :
                driver.execute_script("document.getElementById('event').innerText="+chr(34)+""+chr(34))

                #getdata = driver.execute_script("return document.getElementById('menu').value")

                #新規作成
                if getdata == "shinki" :

                    next_login = p2.new_create(loginUser, name_sel, name_table, name_table2, Shouninsya_sel, Shouninsya_table, Shouninsya_table2, driver)

                    if next_login=="shuuryou":
                        GoTo_continue1=1        ####個人メニュー
                        break
                    elif next_login != "" :
                        GoTo_Start=1
                        break
                    else:
                        GoTo_continue1=1
                        break


                #受信時
                elif getdata == "jyusin" :

                    next_login = p4.Jyushin_syori(loginUser, name_sel, name_table, name_table2, Shouninsya_sel, Shouninsya_table, Shouninsya_table2, driver)
                    if next_login=="shuuryou":
                        GoTo_continue1=1        ####個人メニュー
                        break
                    elif next_login != "" :
                        GoTo_Start=1
                        break
                    else:
                        GoTo_continue1=1
                        break


                #自分の案件
                elif getdata == "jissi" :

                    p5.Create_Table(loginUser, "personal", name_table, driver)

                    GoTo_continue1=1
                    break

                #全部の案件
                elif getdata == "hyou" :
                    p5.Create_Table(loginUser, "all", name_table, driver)

                    GoTo_continue1=1
                    break

                #終了
                elif getdata == "cancel" :
                    driver.close()
                    driver.quit()

                    j.End()

                elif getdata == "modoru" :
##                    driver.close()
##                    driver.quit()
                    next_login = ""
                    GoTo_Start=1
                    break

            j.sleep (10)
##        except:
##            driver.quit()
##            j.End()


if __name__ == '__main__':
    main()


