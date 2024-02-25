
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


    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.edge.service import Service

    profile_path = j.path_join(r"C:\Users",j.getuser(),r"Local\Microsoft\Edge\User Data\Profile 1")
    options = webdriver.EdgeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])
    options.use_chromium = True
    options.add_argument('disable-web-security')
    options.add_argument("--user-data-dir="+ profile_path)
    options.add_argument('--lang=ja')
    ##options.add_argument('--headless')
    options.add_argument("--no-sandbox")
    service = Service(options=options)

    profile_path2 = j.path_join(r"C:\Users",j.getuser(),r"Local\Microsoft\Edge\User Data\Profile 2")
    options2 = webdriver.EdgeOptions()
    options2.add_experimental_option("excludeSwitches", ['enable-automation'])
    options2.use_chromium = True
    options2.add_argument('disable-web-security')
    options2.add_argument("--user-data-dir="+ profile_path2)
    options2.add_argument('--lang=ja')
    ##options.add_argument('--headless')
    options2.add_argument("--no-sandbox")
    service2 = Service(options=options2)


    g.driver0 = webdriver.Edge(service=service,options=options)
    driver = webdriver.Edge(service=service2,options=options2)

    '''
    #ACCESS
    import pyodbc
    import pandas as pd
    '''

    j.get_gamen_size()



    #グローバル変数にはg.を付ける



    driver.get("https://p3.phoneappli.net/front/home")
    j.IE_msg("hohohohO")
    g.driver0.quit()










if __name__ == '__main__':
    main()
