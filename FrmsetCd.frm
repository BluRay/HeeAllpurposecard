VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmsetCd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "制作设置卡"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "FrmsetCd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7815
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "FrmsetCd.frx":030A
      TabIndex        =   20
      Top             =   3720
      Width           =   3135
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6960
      OleObjectBlob   =   "FrmsetCd.frx":0378
      Top             =   4320
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   12
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmsetCd.frx":05AC
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "设置卡参数："
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7335
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmsetCd.frx":0E59
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmsetCd.frx":0EC3
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "FrmsetCd.frx":0F2B
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmsetCd.frx":0F93
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FrmsetCd.frx":0FFD
         TabIndex        =   15
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7095
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "FrmsetCd.frx":1063
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmsetCd.frx":10D1
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   5040
            TabIndex        =   10
            Text            =   "2"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "FrmsetCd.frx":1143
            Left            =   2280
            List            =   "FrmsetCd.frx":1153
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   3240
      OleObjectBlob   =   "FrmsetCd.frx":116F
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmsetCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim rst As Recordset
Dim rst1 As Recordset
Dim SBhao As String
Dim Caiy As Integer
Dim Para(20) As Byte    '21字节
Dim oldpass As String * 4
Dim password(1) As Byte




Private Sub Combo1_Click()
Select Case Combo1.Text
    Case "表一"
    SBhao = "1"
    Case "表二"
    SBhao = "2"
    Case "表三"
    SBhao = "3"
    Case "表四"
    SBhao = "4"
End Select
'填充此表的参数
Set rst = mconn.Execute("select * from WTSdb where wt_no='" + SBhao + "'")
If rst.EOF Or (Trim(rst.Fields("wt_type")) = "") Then
MsgBox "没有设置此表参数，请先设置！"
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
Exit Sub
Else
Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields("wt_type") + "'")
Text2 = rst1.Fields("w_max")
Text3 = rst1.Fields("w_warn1")
Text4 = rst1.Fields("w_warn2")
rst1.Close
Set rst1 = mconn.Execute("select * from WTdb where wt_type='" + rst.Fields("wt_stype") + "'")
Text6 = rst1.Fields("wt_chaiyan")
rst1.Close
End If
rst.Close

End Sub


Private Sub Command1_Click()
On Error GoTo errhandle
If Combo1.Text = "" Then
MsgBox "请选择要设置的表号！"
Exit Sub
End If
'预购量不为0
If Text1.Text = 0 Then
    MsgBox "预购量水量不能为0！"
    Text1.SetFocus
    Exit Sub
End If
'判断IC卡是否准备好
If Not InitICcard Then
    ExitIC
    Exit Sub
End If
st = chk_102(icdev)             '测试是否为合法卡
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
    MsgBox ("核对IC卡密码错！请使用新卡或先回收旧卡！")
    Exit Sub
End If
'********************************************************
'是否开启阶梯水价
'********************************************************


If Val(Trim(Text6)) = 0.1 Then
    Caiy = 1
End If
'写卡
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
''*************写设置卡标志***********************
Para(0) = &H20
st = swr_102_hex(icdev, 0, 18, 1, Para(0))

If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If
'*************************************************
st = asc_hex(Text5, Para(0), 2)     '区码
Para(2) = SBhao                     '表号
Para(3) = Caiy                      '硬参-采样
Para(4) = Val(Text2) \ 256
Para(5) = Val(Text2) Mod 256        '限购量
Para(6) = Val(Text3)                '显警
Para(7) = Val(Text4)                '阀警
Para(8) = Val(Text1)                '预购量-吨
If JTYes Then
Set rst = mconn.Execute("select * from sysJT")
    Para(9) = &H88                      '阶梯水价启用标志-启用
    Para(10) = &H0
    Para(11) = Val(rst.Fields("nian1"))                     '阶梯值
    Para(12) = &H0
    Para(13) = Val(rst.Fields("nian2"))                     '阶梯值
    
    Dim Jia1 As Integer, Jia2 As Integer, Jia3 As Integer
    Jia1 = Val(rst.Fields("Jia1")) * 100
    Jia2 = Val(rst.Fields("Jia2")) * 100
    Jia3 = Val(rst.Fields("Jia3")) * 100
    Para(14) = Jia1 Mod 100
    Para(15) = Jia1 \ 100
    
    Para(16) = Jia2 Mod 100
    Para(17) = Jia2 \ 100
    
    Para(18) = Jia3 Mod 100
    Para(19) = Jia3 \ 100
    
    Para(20) = &H0
    For i = 0 To 19
    Para(20) = Para(20) Xor Para(i)
    Next i
rst.Close
Else
    Para(9) = &HFF                      '阶梯水价启用标志-不启用
    Para(10) = &HFF
    Para(11) = &HFF                     '阶梯值
    Para(12) = &HFF
    Para(13) = &HFF
    Para(14) = &HFF
    Para(15) = &HFF
    Para(16) = &HFF
    Para(17) = &HFF
    Para(18) = &HFF
    Para(19) = &HFF
    Para(20) = &H0                      '校验
    For i = 0 To 19
    Para(20) = Para(20) Xor Para(i)
    Next i
End If
'**************擦除地址******************************
st = ser_102(icdev, 2, 0, 22)
If st < 0 Then
    MsgBox ("擦卡失败！")
    Exit Sub
End If

Screen.MousePointer = vbHourglass

st = swr_102_hex(icdev, 2, 1, 21, Para(0))
If st < 0 Then
  MsgBox ("写卡失败！！")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'********************************************************
'********************************************************

'2区读保护位清零
'*************读保护位清0,核对密码前不能对应用区2进行读操作*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("读保护位清零出错！")
  Exit Sub
End If

'*************更改密码************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("更改卡密码出错！")
    Exit Sub
End If

MsgBox "设置卡制作成功！"
Screen.MousePointer = vbDefault
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
'**********填充地区代码********************
Set rst = mconn.Execute("select area from sysdate")
Text5 = rst.Fields(0)
rst.Close
'**********当前是否开通阶梯水价**********
If JTYes Then
SkinLabel9.Caption = "已启用价梯水价！"
SkinLabel3.Caption = "预购水量：(元)"
Else
SkinLabel9.Caption = "未启用价梯水价！"
SkinLabel3.Caption = "预购水量：(吨)"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 27 Then   'ESC键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 27 Then   'ESC键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 27 Then   'ESC键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 27 Then   'ESC键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub

