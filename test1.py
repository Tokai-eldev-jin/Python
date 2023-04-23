#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      jinkawa-san
#
# Created:     23/04/2023
# Copyright:   (c) jinkawa-san 2023
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def main():
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")


    #すでにExcelが起動されている場合はそのタスクが使われる
    #エラー終了するとタスクは残ります
    xl = win32com.client.Dispatch("Excel.Application")
    #動いている様子を見てみる
    xl.Visible = True

    #ブック追加
    wb = xl.Workbooks.Add()

    #シート追加
    ws = wb.Worksheets.Add()

    #シート名変更
    ws.Name = "練習1"

    #不要シート削除
    xl.DisplayAlerts = False #無くても良いようだが
    for sh in wb.Worksheets:
        if sh.Name != ws.Name:
            sh.Delete()

    # セルに値を設定
    for i, cell in enumerate(ws.Range("A1:B5")):
        cell.Value = i

    #セル値を2次元配列で取得：タプルで取得される
    ary1 = ws.Range("A1:B5").Value
    print(f"{ary1=}")

    #セルに入れるために2次元タプルを2次元リストに変換
    ary2 = [[c for c in r] for r in ary1]
    print(f"{ary2=}")

    #2次元配列の値を2倍
    for r, row in enumerate(ary2):
        for c, col in enumerate(row):
            ary2[r][c] = col * 2

    #2次元配列をセルに入れる
    ws.Range("A1:B5").Value = ary2

    #Rangeのコピー
    ws.Range("A1:B5").Copy(Destination=ws.Range("D1"))

    #セルの削除
    ws.Columns(5).Delete()

    #セルの削除
    ws.Range(ws.Cells(2,4),ws.Cells(4,4)).Delete(Shift=-4162) #-4162=xlUp

    #罫線設定
    ws.Range("A1:B5").Borders.LineStyle = 1

    #表示形式
    ws.Range("B1").NumberformatLocal = "0.0%"

    #フォント
    ws.Range("A1").CurrentRegion.Font.Bold = True
    ws.Range("A1").CurrentRegion.Font.Size = 14
    ws.Range("A1").CurrentRegion.Font.Color = 0x0000FF #赤

    #その他のRangeのプロパティ・メソッドの確認
    ws.Rows(3).Hidden = True
    ws.Range("A1:B5").EntireColumn.AutoFit()

    #これ以降は出力シートを変更
    ws = wb.Worksheets.Add()
    ws.Name = "練習2"

    #以降の結果をセルに出力するので：セル出力には2次元配列が必要
    titles = (("Sum",),
              ("CountIf",),
              ("Match",),
              ("VLookup",),
              ("Range(""A1"")",),
              ("Range(""A1:B5"")",),
              ("Range(Cells,Cells)",),
              ("UsedRange",),
              ("CurrentRegion",),
              ("SpecialCells",),
              ("End(xlUp)",),
              ("Offset(1, 1)",),
              ("Offset(2, 2)",),
              ("Offset(3, 3)",),
              ("Resize(2, 2)",),
              ("Resize(3, 3)",),
              ("Address(0, 0)",),
              ("InputBox",))
    ws.Range("A1:A18").Value = titles

    #WorksheetFunction
    ws.Range("B1").Value = xl.WorksheetFunction.Sum(ws.Range("A1:A5"), ws.Range("B1:B5"))
    ws.Range("B2").Value = xl.WorksheetFunction.CountIf(ws.Range("A1:B5"),">10")
    try: #検索値が無い場合はエラーになるので
        ws.Range("B3").Value = xl.WorksheetFunction.Match("Match", ws.Range("A:A"), 0)
    except Exception as e:
        #タプルの入れ子かつ数値文字混在なのでちょっと面倒です。
        ws.Range("B3").Value = ",".join([str(x) for x in e.args[2]])
    try:
        ws.Range("B4").Value = xl.WorksheetFunction.VLookup("VLookup", ws.Range("A:A"), 1, False)
    except Exception as e:
        ws.Range("B4").Value = ",".join([str(x) for x in e.args[2]])

    #Offset：1ずれている?、承知して使うのは怖い気がする
    ws.Range("B12").Value = ws.Range("A1").Offset(1, 1).Address
    ws.Range("B13").Value = ws.Range("A1").Offset(2, 2).Address
    ws.Range("B14").Value = ws.Range("A1").Offset(3, 3).Address

    #Resize：これはおかしいので使えない
    ws.Range("B15").Value = ws.Range("A1").Resize(2, 2).Address
    ws.Range("B16").Value = ws.Range("A1").Resize(3, 3).Address

    #Addressに引数が指定できない?：'str' object is not callable
    try:
        ws.Range("B17").Value = ws.Range("B27").Value = ws.Range("A1").Address(False, False)
    except Exception as e:
        ws.Range("B17").Value = str(e)

    #InputBox：
    ws.Range("B18").Value = xl.Inputbox("入力してください。")

    #ブックを保存：相対パスはExcel.Applicationから見た位置になります。
    import os
    wb.SaveAs(os.getcwd() + "/test.xlsx") #.pyと同ディレクトリ
    wb.Close(SaveChanges=True)

    #他のブックが開いていなければExcelを終了
    if xl.Workbooks.Count == 0:
        xl.quit()

if __name__ == '__main__':
    main()
