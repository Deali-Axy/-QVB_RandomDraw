VERSION 5.00
Begin VB.Form Frm_Begin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "QVB 随机抽签"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "微软雅黑"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Btn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "抽签"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "确定"
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
         Name            =   "微软雅黑"
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
         Name            =   "宋体"
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
      Caption         =   "抽签范围："
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
'工程名称：QVB示例-随机抽签
'作者：Deali-Axy
'时间：2016.1.30
'简介：这是一个使用QVB库开发的例子，简单运用了一些QVB的功能，
'           短短几行代码就实现了一个基本的抽签器，QVB库封装的功能
'           着实方便了程序的开发
'QVB简介：QVB是我在2013年开发的一个API封装库，封装了一些常用
'               的简单功能，可以方便VB程序的开发
'作者邮箱：deali@live.com QQ：1875615476 欢迎技术交流

Dim QVB As New QVB    '创建QVB对象

'抽签范围
Dim UpBound As Integer
Dim LrBound As Integer


Private Sub Btn_Click()
    Dim intTmp As Long
    intTmp = QVB_RandomInt(LrBound, UpBound)    '获取随机数
    Lbl = Trim(Str(intTmp))
End Sub

Private Sub Btn_OK_Click()

'数据校验
    If Len(Txt_UpBound) = 0 Then
        MsgBox "没输入完整"
        Exit Sub
    End If
    If Len(Txt_LrBound) = 0 Then
        MsgBox "没输入完整"
        Exit Sub
    End If

    UpBound = Val(Txt_UpBound)
    LrBound = Val(Txt_LrBound)
End Sub

Private Sub Form_Load()
    QVB.QWindow.AnimateForm Me, aLoad, eCurtonHorizontal, 40    '设置窗口动画
    QVB.QWindow.AlwaysOnTop Me.hWnd, True    '设置窗口置顶
    QVB.QWindow.Transparent Me.hWnd, 10    '设置窗口透明度

'读取配置
    QVB.QSystem.QIniSettings.INIFile = App.Path & "\config.ini"
    UpBound = Val(QVB.QSystem.QIniSettings.GetText("QVB RandomDraw", "UpBound"))
    LrBound = Val(QVB.QSystem.QIniSettings.GetText("QVB RandomDraw", "LrBound"))
    If UpBound = 0 Then UpBound = 100    '防止读取不到配置后出错
    Txt_UpBound = Trim(Str(UpBound))
    Txt_LrBound = Trim(Str(LrBound))
End Sub

Private Sub Form_Unload(Cancel As Integer)
'保存配置
    QVB.QSystem.QIniSettings.WriteText "QVB RandomDraw", "UpBound", Str(UpBound)
    QVB.QSystem.QIniSettings.WriteText "QVB RandomDraw", "LrBound", Str(LrBound)

    QVB.QWindow.AnimateForm Me, aUnload, eFoldOut, 40    '窗口关闭动画
End Sub

Private Sub Lbl_Author_Click()
    QVB.QNetwork.URL "http://weibo.com/dealiaxy"    '调用网址
End Sub
