
import jin as j
import global_value as g


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
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.common.by import By



    options = webdriver.EdgeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])

    profile_path = j.path_join(r"C:\Users",j.getuser(),r"Local\Microsoft\Edge\User Data\Profile 1")
    options.use_chromium = True
    options.add_argument('disable-web-security')
    options.add_argument("--user-data-dir="+ profile_path)
    options.add_argument('--lang=ja')
    ##options.add_argument('--headless')
    options.add_argument("--no-sandbox")


    if j.File_system("fs_exists",j.path_join(r"C:\Users",j.getuser(),
                                r"AppData\Local\SeleniumBasic\edgedriver.exe")) == True:
        g.edgeDriver = j.path_join(r"C:\Users",j.getuser(),r"AppData\Local\SeleniumBasic\edgedriver.exe")
    else:
        g.edgeDriver = r"C:\Program Files\SeleniumBasic\edgedriver.exe"

    service = Service(executable_path=g.edgeDriver,options=options)



    ##fukidasiやIE_msg用chromeの立ち上げ
##    fuki_id=0
##    g.driver0 = webdriver.Edge(service=service, options=options)

    ####################Edge立ち上げ####################
##    driver = webdriver.Edge(service=service, options=options)
##    driver.set_page_load_timeout(5)
##    driver.maximize_window()
##
##    driver.get("https://p3.phoneappli.net/front/home")
##    j.BusyWait(driver)
##    j.sleep(10)
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
