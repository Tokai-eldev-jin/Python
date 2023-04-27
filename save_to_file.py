#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      A0044
#
# Created:     26/04/2023
# Copyright:   (c) A0044 2023
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    pass

import jin as j
import global_value as g

def save_to_file(driver0):


    xlContinuous=1

    j.gettime(0)
    today_yy2 = g.G_TIME_YY4
    today_mm2 = g.G_TIME_MM2
    today_dd2 = g.G_TIME_DD2
    today_hh2 = g.G_TIME_HH2
    today_nn2 = g.G_TIME_NN2
    today_ss2 = g.G_TIME_SS2

    ex,wb = j.ex_OPEN()
    ex.Visible = True
    ex.ScreenUpdating = True


    savepath = j.path_join(g.CUR_DIR , "file" , today_yy2 + today_mm2 + today_dd2 + "_" + today_hh2 + today_nn2 + ".xlsx")

    g.token_str = j.IE_gettable(driver0, 0)
    gyou = int(j.token("_", g.token_str))
    retu = int(j.token("_", g.token_str))


##    wb = ex.Workbooks.Add
    ws = wb.Sheets(1)

    for i in range( 0 , gyou):
        for ii in range( 0 , retu):
            ws.Range("A1").GetOffset(i, ii).Value = g.IE_table[str(i) + "_" + str(ii)]

    ex.ScreenUpdating = False
    for i in range(1,retu+1):
        for ii in range(1,gyou+1):
            ws.Cells(ii, i).Borders.LineStyle = xlContinuous

    ws.Cells.Select()
    ws.Cells.EntireColumn.AutoFit()

##    ws.Range(ws.Cells(1, 1), ws.Cells(1, retu)).EntireColumn.AutoFit
##    ws.Range(ws.Cells(1, 1), ws.Cells(gyou, retu)).Borders(7).LineStyle = xlContinuous

    ex.DisplayAlerts = False
    wb.SaveAs(Filename:=savepath)
    ex.DisplayAlerts = True

##    j.exec_apli(j.path_join(g.CUR_DIR,"click1.vbs"))
    j.sleep (1000)
    ex.ScreenUpdating = True


    ##Print_PDF(ws)







##def Print_PDF(ws):
##
##    j.gettime(0)
##
##    ""プリンターをPDFに変更する
##    doscmd("rundll32.exe printui.dll,PrintUIEntry /y /n " & Chr(34) & "Microsoft Print to PDF" & Chr(34))
##
##
##    With ws.PageSetup
##        .PaperSize = xlPaperA3
##        .Orientation = xlLandscape
##        .Zoom = False
##        .FitToPagesWide = 1
##    End With
##
##    ws.PrintOut
##
##    'プリンターの設定を元に戻す
##    doscmd "rundll32.exe printui.dll,PrintUIEntry /y /n " & Chr(34) & "\\hnps-sv04\TRJ-00037" & Chr(34)
##
##
##    exec CUR_DIR & "\file\情報機器管理_" & G_TIME_YY4 & G_TIME_MM2 & G_TIME_DD2 & ".pdf"
##





if __name__ == '__main__':
    main()
