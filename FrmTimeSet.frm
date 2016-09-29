VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmTimeSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "水表时间设置"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7080
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "FrmTimeSet.frx":0000
      TabIndex        =   11
      Top             =   2880
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1920
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmTimeSet.frx":00E6
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6600
      OleObjectBlob   =   "FrmTimeSet.frx":0A9A
      Top             =   2520
   End
   Begin VB.Frame Frame1 
      Caption         =   "设置水表时间："
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   5400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   4200
         TabIndex        =   4
         Text            =   "Combo4"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   2520
         TabIndex        =   3
         Text            =   "Combo3"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmTimeSet.frx":0CCE
         TabIndex        =   6
         Top             =   540
         Width           =   5295
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   2760
      OleObjectBlob   =   "FrmTimeSet.frx":0DE8
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim password(1) As Byte
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
oldpass = "1b6c"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("核对IC卡密码错")
    Exit Sub
End If
'**************先判断卡型******************
'读卡信息
    st = srd_102_hex(icdev, 0, 18, 1, RdType)
    If RdType <> &H10 Then
    MsgBox "此卡不是用户卡！请正确插入用户卡！"
    Exit Sub
    End If
Dim Para(5) As Byte
    Para(0) = &H35                          '设时间标志
    Para(1) = Val(Right(Combo1, 2))         '年
    Para(2) = Val(Combo2)                   '月
    Para(3) = Val(Combo3)                   '日
    Para(4) = Val(Combo4)                   '时
    Para(5) = Val(Text1)                    '分
st = ser_102(icdev, 2, 54, 6)
If st < 0 Then
    MsgBox ("擦卡出错！！")
    Exit Sub
End If
st = swr_102_hex(icdev, 2, 54, 6, Para(0))
If st < 0 Then
MsgBox "写卡出错！"
Exit Sub
End If

MsgBox "做卡成功！！"
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

