Attribute VB_Name = "Mod_QVBAssist"
Option Explicit
'QVB������ĸ���ģ��

Dim QVB_Rnd As New QVB_Random

'��ȡ�������
Function QVB_RandomInt(intLrBound As Integer, intUpBound As Integer) As Long
    QVB_RandomInt = QVB_Rnd.RandomInt(intLrBound, intUpBound)
End Function
