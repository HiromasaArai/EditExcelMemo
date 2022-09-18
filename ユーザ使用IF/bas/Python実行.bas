Attribute VB_Name = "Python実行"
Option Explicit

Sub 新規メモ作成()
    Call GoPython("新規メモ作成.exe")
End Sub


Sub メモ更新()
    Call GoPython("メモ更新.exe")
End Sub


Sub 成果物をPDF化()
    Call GoPython("成果物をPDF化.exe")
End Sub


Sub 索引検索()
    Call GoPython("索引検索.exe")
End Sub


Sub 別名登録()
    Call GoPython("別名登録.exe")
End Sub
