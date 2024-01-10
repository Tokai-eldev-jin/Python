
import jin as j
import global_value as g



def AAA():
    hnd=j.getid("all_win")
    for i in hnd:
        print(i)



def main():


    import os
    import sys
    import time

    def get_CUR_DIR():
        import sys
        import os

        path_current_dir = os.path.dirname(sys.argv[0])

        #print(sys.argv[0])
        #print(path_current_dir)
        return  path_current_dir

    g.CUR_DIR=get_CUR_DIR()

    #print(sys.argv[0])#起動パラメータ


    #web
    #必要なライブラリをインポート
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from msedge.selenium_tools import Edge, EdgeOptions

    options = webdriver.EdgeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])
    profile_path = j.path_join(r"C:\Users",j.getuser(),r"Local\Microsoft\Edge\User Data\Profile 1")
    options.use_chromium = True
    options.add_argument('disable-web-security')
    options.add_argument("--user-data-dir="+ profile_path)
    options.add_argument('--lang=ja')
    ##options.add_argument('--headless')
    options.add_argument("--no-sandbox")

    options2 = webdriver.EdgeOptions()
    options2.add_experimental_option("excludeSwitches", ['enable-automation'])
    profile_path2 = j.path_join(r"C:\Users",j.getuser(),r"Local\Microsoft\Edge\User Data\Profile 2")
    options2.use_chromium = True
    options2.add_argument('disable-web-security')
    options2.add_argument("--user-data-dir="+ profile_path2)
    options2.add_argument('--lang=ja')
    ##options.add_argument('--headless')
    options2.add_argument("--no-sandbox")





    ##fukidasiやIE_msg用chromeの立ち上げ
    fuki_id=0
    g.driver0 = webdriver.Edge(options=options)


    ##j.Thread(AAA)
    j.ctrlwin(j.getid("新しいタブ - 個人 - Microsoft​ Edge","Chrome_WidgetWin_1"),"fs_min")


    ####################Edge立ち上げ####################
    driver = webdriver.Edge(options=options2)
    driver.set_page_load_timeout(5)
    driver.maximize_window()

    driver.get("https://www.yahoo.co.jp/")
    j.BusyWait(driver)
    j.sleep(10)
    ####################Edge立ち上げ####################


    '''
    #ACCESS
    import pyodbc
    import pandas as pd
    '''

    j.get_gamen_size()



    #グローバル変数にはg.を付ける






    j.IE_msg("hohohohO")










if __name__ == '__main__':
    main()
