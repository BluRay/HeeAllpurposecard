VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmZero 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "制作水表清零卡"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "FrmZero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6120
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "FrmZero.frx":6F0C2
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1440
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmZero.frx":6F2F6
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   2160
      OleObjectBlob   =   "FrmZero.frx":6FCAA
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   855
      Left            =   360
      OleObjectBlob   =   "FrmZero.frx":6FD0B
      TabIndex        =   3
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "FrmZero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdType(0) As Byte
Dim i As Integer
Dim Para(2) As Byte
Dim oldpass As String * 4
Dim password(1) As Byte


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
    MsgBox ("核对IC卡密码错！必须使用新卡或先清卡。")
    Exit Sub
End If
''如果是用户卡的话，先提示
'st = srd_102_hex(icdev, 2, 3, 1, RdType(0))
'If RdType(0) = &H10 Then
'MsgBox "此卡为用户卡，不能用作他用。如确认此卡已作废，请先清卡！"
'Exit Sub
'End If

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
''*************写清零卡标志***********************
Para(0) = &H30
st = swr_102_hex(icdev, 0, 18, 1, Para(0))

If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If


''**************擦除2区******************************
'st = ser_102(icdev, 2, 1, 3)
'If st < 0 Then
'    MsgBox ("擦卡出错！！")
'    Exit Sub
'End If
''*************写清零卡标志***********************
'Dim rst As Recordset
'Set rst = mconn.Execute("select area from Sysdate")
'st = asc_hex(rst.Fields(0), Para(0), 2)
'Para(2) = &H40
'rst.Close
'st = swr_102_hex(icdev, 2, 1, 3, Para(0))
'
'If st < 0 Then
'MsgBox "写卡失败！"
'Exit Sub
'End If
'*************更改密码************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("更改卡密码出错！")
    Exit Sub
Else
    MsgBox ("制作清零卡成功！！")
End If

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

End Sub
