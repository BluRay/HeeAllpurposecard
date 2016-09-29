VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmChuShiHua 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "制作初始化卡"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "FrmChuShiHua.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7245
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Caption         =   "调整水表时间"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   820
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "                "
      Enabled         =   0   'False
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   6855
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   2520
         TabIndex        =   8
         Text            =   "Combo3"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   4200
         TabIndex        =   7
         Text            =   "Combo4"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   5400
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmChuShiHua.frx":030A
         TabIndex        =   11
         Top             =   540
         Width           =   5295
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5880
      OleObjectBlob   =   "FrmChuShiHua.frx":0424
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      Height          =   715
      Left            =   1680
      ScaleHeight     =   660
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmChuShiHua.frx":0658
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "FrmChuShiHua.frx":100C
      TabIndex        =   2
      Top             =   3000
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   2760
      OleObjectBlob   =   "FrmChuShiHua.frx":109E
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmChuShiHua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdType(0) As Byte
Dim i As Integer
Dim Para(7) As Byte
Dim oldpass As String * 4
Dim password(1) As Byte

Private Sub Check1_Click()
If Check1.Value Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
On Error GoTo errhandle
'判断IC卡是否准备好
If Not InitICcard Then
    ExitIC
    Exit Sub
End If
'测试是否为合法卡
st = chk_102(icdev)
If st <> 0 Then
    MsgBox ("不是合法的IC卡！请检查。")
    Exit Sub
End If
'***************核对密码f0f0***************************
'password(0) = &HF0
'password(1) = &HF0
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("核对IC卡密码错")
    Exit Sub
End If
'**************擦除0区******************************
st = ser_102(icdev, 0, 18, 5)
If st < 0 Then
    MsgBox ("擦卡出错！！")
    Exit Sub
End If
'写本系统卡标志
Para(0) = &H98
st = swr_102_hex(icdev, 0, 21, 1, Para(0))
If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If
''*************写卡标志***********************
Para(0) = &H50
st = swr_102_hex(icdev, 0, 18, 1, Para(0))
If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If


'**************擦除2区******************************
st = ser_102(icdev, 2, 21, 8)
If st < 0 Then
    MsgBox ("擦卡出错！！")
    Exit Sub
End If
'**************写时间参数******************************
If Check1.Value Then                    '要设置时间
    Para(0) = Val(Right(Combo1, 2))         '年
    Para(1) = Val(Combo2)                   '月
    Para(2) = Val(Combo3)                   '日
    Para(3) = Val(Combo4)                   '时
    Para(4) = Val(Text1)                    '分
    Para(5) = &H0                           '秒
    Para(6) = &H88                           '设时间标志
    Para(7) = Para(0) Xor Para(1) Xor Para(2) Xor Para(3) Xor Para(4) Xor Para(5) Xor Para(6)                       '校验
Else
    For i = 0 To 7
    Para(i) = &H0
    Next i
End If


st = swr_102_hex(icdev, 2, 21, 8, Para(0))
If st < 0 Then
MsgBox "写卡出错！"
Exit Sub
End If

'*************更改密码************************
'2区读保护位清零
'*************读保护位清0,核对密码前不能对应用区2进行读操作*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("读保护位清零出错！")
  Exit Sub
End If
'***************************************************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("更改卡密码出错！")
    Exit Sub
End If
MsgBox "初始化卡制作成功！"


ExitIC
Unload Me
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'***************填充日期******************
Dim i As Integer
Combo1 = Year(Now)
Combo2 = Month(Now)
Combo3 = Day(Now)
Combo1.AddItem (Year(Now) + 1)
Combo2.AddItem (Month(Now) + 1)
For i = 1 To 31
Combo3.AddItem i
Next i
Combo4 = Hour(Now)
For i = 0 To 23
Combo4.AddItem i
Next i
Text1 = Minute(Now)

End Sub
