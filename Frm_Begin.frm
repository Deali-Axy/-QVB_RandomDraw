VERSION 5.00
Begin VB.Form Frm_Begin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "QVB �����ǩ"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Btn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ǩ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Btn_OK 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Txt_UpBound 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "100"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Txt_LrBound 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Lbl_Author 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Deali-Axy"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1560
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   200.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǩ��Χ��"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "Frm_Begin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�������ƣ�QVBʾ��-�����ǩ
'���ߣ�Deali-Axy
'ʱ�䣺2016.1.30
'��飺����һ��ʹ��QVB�⿪�������ӣ���������һЩQVB�Ĺ��ܣ�
'           �̶̼��д����ʵ����һ�������ĳ�ǩ����QVB���װ�Ĺ���
'           ��ʵ�����˳���Ŀ���
'QVB��飺QVB������2013�꿪����һ��API��װ�⣬��װ��һЩ����
'               �ļ򵥹��ܣ����Է���VB����Ŀ���
'�������䣺deali@live.com QQ��1875615476 ��ӭ��������

Dim QVB As New QVB    '����QVB����

'��ǩ��Χ
Dim UpBound As Integer
Dim LrBound As Integer


Private Sub Btn_Click()
    Dim intTmp As Long
    intTmp = QVB_RandomInt(LrBound, UpBound)    '��ȡ�����
    Lbl = Trim(Str(intTmp))
End Sub

Private Sub Btn_OK_Click()

'����У��
    If Len(Txt_UpBound) = 0 Then
        MsgBox "û��������"
        Exit Sub
    End If
    If Len(Txt_LrBound) = 0 Then
        MsgBox "û��������"
        Exit Sub
    End If

    UpBound = Val(Txt_UpBound)
    LrBound = Val(Txt_LrBound)
End Sub

Private Sub Form_Load()
    QVB.QWindow.AnimateForm Me, aLoad, eCurtonHorizontal, 40    '���ô��ڶ���
    QVB.QWindow.AlwaysOnTop Me.hWnd, True    '���ô����ö�
    QVB.QWindow.Transparent Me.hWnd, 10    '���ô���͸����

'��ȡ����
    QVB.QSystem.QIniSettings.INIFile = App.Path & "\config.ini"
    UpBound = Val(QVB.QSystem.QIniSettings.GetText("QVB RandomDraw", "UpBound"))
    LrBound = Val(QVB.QSystem.QIniSettings.GetText("QVB RandomDraw", "LrBound"))
    If UpBound = 0 Then UpBound = 100    '��ֹ��ȡ�������ú����
    Txt_UpBound = Trim(Str(UpBound))
    Txt_LrBound = Trim(Str(LrBound))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'��������
    QVB.QSystem.QIniSettings.WriteText "QVB RandomDraw", "UpBound", Str(UpBound)
    QVB.QSystem.QIniSettings.WriteText "QVB RandomDraw", "LrBound", Str(LrBound)

    QVB.QWindow.AnimateForm Me, aUnload, eFoldOut, 40    '���ڹرն���
End Sub

Private Sub Lbl_Author_Click()
    QVB.QNetwork.URL "http://weibo.com/dealiaxy"    '������ַ
End Sub
