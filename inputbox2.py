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


def inputbox2(msg,default1=1,default2=32767):

    import jin as j
    import global_value as g
    from selenium import webdriver
##    import chromedriver_autoinstaller

##    chromedriver_autoinstaller.install()
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ['enable-automation'])

    web_str = "<html>"
    web_str = web_str + "<head>"
    web_str = web_str + "<title>入力してください。</title>"
    web_str = web_str + "<script language='JavaScript'>"
    web_str = web_str + "function eventA(getdata){"
    web_str = web_str + "   if(getdata=='atai1'){"
    web_str = web_str + "       document.getElementById('atai2').select()"
    web_str = web_str + "   }else if(getdata=='atai2'){"
    web_str = web_str + "      document.getElementById('btn1').focus()"
    web_str = web_str + "   }"
    web_str = web_str + "}"
    web_str = web_str + "</script>"
    web_str = web_str + "</head>"
    web_str = web_str + "<body width='30' height='60'>"
    web_str = web_str + "<span id='alete'><font size='4'>" + msg + "</font></span>"
    web_str = web_str + "<br><br><input id='atai1' type='text' size='3' onchange='eventA(this.id)' value = '" + str(default1) + "'>"
    web_str = web_str + "　-　<input id='atai2' type='text' size='3' onchange='eventA(this.id)' value = '" + str(default2) + "'>"
    web_str = web_str + "<br><br><br><input id='btn1' type='button' value = 'ok' onClick='value=" + chr(34) + "-" + chr(34) + "'>"
    web_str = web_str + "　　　<input id='btn2' type='button' value = 'cancel' onClick='value=" + chr(34) + "-" + chr(34) + "'>"
    web_str = web_str + "<span id='eventAA'  style='display:none;'></span></body>"
    web_str = web_str + "</html>"





    f1 = j.fopen(j.path_join(g.CUR_DIR , "Temp.html"), "f_write")
    j.fput(f1, web_str)
    j.fclose(f1)


    driver0 = webdriver.Chrome(options=options)
    ##    driver.maximize_window()
    driver0.set_window_size(g.g_screen_w * 4 / 5, 400)
    driver0.get(j.path_join(g.CUR_DIR,"Temp.html"))
    j.BusyWait(driver0)
    j.sleep(100)
    driver0.refresh()


    driver0.execute_script("document.getElementById('atai1').select()")

    ##try
    input_cyc = 0
    while True:

        getdata = driver0.execute_script("return document.getElementById('btn1').value")

        if getdata == "-" :
            inputbox2 = driver0.execute_script("return document.getElementById('atai1').value") + "-" + driver0.execute_script("return document.getElementById('atai2').value")
            driver0.close()
            return inputbox2

        getdata = driver0.execute_script("return document.getElementById('btn2').value")

        if getdata == "-" :
            inputbox2 = False

            driver0.close()
            return inputbox2




    ##exceppt

    return False






if __name__ == '__main__':
    main()
