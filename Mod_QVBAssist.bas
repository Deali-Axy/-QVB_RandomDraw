Attribute VB_Name = "Mod_QVBAssist"
Option Explicit
'QVB开发库的辅助模块

Dim QVB_Rnd As New QVB_Random

'获取随机整数
Function QVB_RandomInt(intLrBound As Integer, intUpBound As Integer) As Long
    QVB_RandomInt = QVB_Rnd.RandomInt(intLrBound, intUpBound)
End Function
