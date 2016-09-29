VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FRMYHkhD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "售电开户"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9270
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   29
      Text            =   "0"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   9015
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FRMYHkhD.frx":0000
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   270
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FRMYHkhD.frx":0068
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "FRMYHkhD.frx":00D0
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   270
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FRMYHkhD.frx":013C
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   270
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   6240
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FRMYHkhD.frx":01A8
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FRMYHkhD.frx":0214
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   4560
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMYHkhD.frx":027C
      Top             =   4920
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2880
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FRMYHkhD.frx":04B0
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "用户信息："
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9015
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   5280
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "按身份证号购买："
         Height          =   255
         Left            =   3480
         TabIndex        =   39
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1800
         TabIndex        =   0
         ToolTipText     =   "提示：用户编号前的0可省略"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按用户编号购买："
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "确 定"
         Height          =   375
         Left            =   7440
         TabIndex        =   1
         Top             =   160
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   8775
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   160
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   400
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   160
            Width           =   4095
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   400
            Width           =   4095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "FRMYHkhD.frx":0D5D
            TabIndex        =   10
            Top             =   405
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "FRMYHkhD.frx":0DD1
            TabIndex        =   11
            Top             =   165
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FRMYHkhD.frx":0E39
            TabIndex        =   12
            Top             =   405
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FRMYHkhD.frx":0EA1
            TabIndex        =   13
            Top             =   165
            Width           =   975
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   3600
      OleObjectBlob   =   "FRMYHkhD.frx":0F09
      TabIndex        =   14
      Top             =   120
      Width           =   3135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "FRMYHkhD.frx":0F6A
      TabIndex        =   28
      Top             =   3600
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "FRMYHkhD.frx":0FD2
      TabIndex        =   32
      Top             =   3840
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   375
      Left            =   6240
      OleObjectBlob   =   "FRMYHkhD.frx":103A
      TabIndex        =   33
      Top             =   4080
      Width           =   3015
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   6360
      OleObjectBlob   =   "FRMYHkhD.frx":1097
      TabIndex        =   34
      Top             =   3600
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FRMYHkhD.frx":1101
      TabIndex        =   35
      Top             =   3840
      Width           =   3495
   End
End
Attribute VB_Name = "FRMYHkhD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim rst As Recordset, rst1 As Recordset
Dim oldpass As String * 4
Dim password(1) As Byte
Dim Para(25) As Byte   '参数数组，共26字节

Private Sub Combo1_Click()
'根据用电类型填充详细信息
Set rst = mconn.Execute("select * from WTDdb where Ds_name='" + Combo1 + "'")
Text2 = rst.Fields("Ds_price")
Text3 = rst.Fields("Ds_gznum")
Text4 = rst.Fields("Ds_tz")
rst.Close
End Sub

Private Sub Command1_Click()
'开户购电
On Error GoTo errhandle
'购电量不能为零
If Val(Text8) = 0 Then
    MsgBox "您还没有购电，请购电！"
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
'***************考虑可能先进行水表开户，比两次密码***************
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    oldpass = "1b6c"
    st = asc_hex(oldpass, password(0), 2)
    st = csc_102(icdev, 2, password(0))
    If st < 0 Then
    MsgBox ("核对IC卡密码错！")
    Exit Sub
    End If
End If
'**************生成用户卡参数******************************
Dim TempGD As String        '购电量
Dim TempGZ As String        '过载次数
Dim TempTZ As String        '透支量
'用户编号00000001
'Tempid = FormatString(Text15, 8)
'擦除0区
st = ser_102(icdev, 0, 2, 8)
If st < 0 Then
    MsgBox "擦卡失败！！"
    Exit Sub
End If
'写用户编号************************
Para(0) = &H0
Para(1) = &H0
Call ToBCD(Left(Right(Text15, 4), 2), Para(2))
Call ToBCD(Right(Text15, 2), Para(3))
'开户标志**************************
Para(4) = &H41
st = swr_102_hex(icdev, 0, 2, 5, Para(0))
If st < 0 Then
  MsgBox ("写卡出错！！")
  Exit Sub
End If
''擦除0区23写bf
'st = ser_102(icdev, 0, 2, 23)
'If st < 0 Then
'    MsgBox "擦卡失败！！"
'    Exit Sub
'End If
'Dim TempBf As Byte
'TempBf = &HBF
'st = swr_102_hex(icdev, 0, 22, 1, TempBf)
'If st < 0 Then
'  MsgBox ("写卡出错！！")
'  Exit Sub
'End If



'擦除1区***************************
st = ser_102(icdev, 1, 0, 22)
If st < 0 Then
    MsgBox "擦卡失败！！"
    Exit Sub
End If

'卡密码****************************
Para(5) = &HC2
Para(6) = &HA9
Dim Apass As String
Set rst = mconn.Execute("select Apass from Sysdate")
Apass = rst.Fields(0)
'Call ToBCD(Val(Left(Apass, 2)), Para(7))
'Call ToBCD(Val(Right(Apass, 2)), Para(8))
st = asc_hex(Apass, Para(7), 2)
If st < 0 Then
    MsgBox ("读取卡密码错")
    Exit Sub
End If
rst.Close

'Para(7) = &H36
'Para(8) = &H10
'购电量
TempGD = FormatString(Val(Text1), 4)
Call ToBCD(Left(TempGD, 2), Para(9))
Call ToBCD(Right(TempGD, 2), Para(10))
'过载次数
TempGZ = FormatString(Val(Text3), 2)
Call ToBCD(TempGZ, Para(11))
'透支量
TempTZ = FormatString(Val(Text4), 2)
Call ToBCD(TempTZ, Para(12))
For i = 13 To 17
Para(i) = &H0
Next i
Para(18) = &H1
For i = 19 To 24
Para(i) = &H0
Next i

st = swr_102_hex(icdev, 1, 2, 20, Para(5))
If st < 0 Then
  MsgBox ("写卡出错！！")
  Exit Sub
End If
'1区读保护位清零
'*************读保护位清0,核对密码前不能对应用区1进行读操作*****
st = clrrd_102(icdev, 1)
If st < 0 Then
  MsgBox ("读保护位清零出错！")
  Exit Sub
End If
'*************更改1区擦除密码为2cc1067d9435************************
Dim pass(6) As Byte
pass(0) = &H2C
pass(1) = &HC1
pass(2) = &H6
pass(3) = &H7D
pass(4) = &H94
pass(5) = &H35
st = wesc_102(icdev, 1, 6, pass(0))
If st < 0 Then
    MsgBox ("更改卡1区擦除密码出错！")
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

'保存开户信息到数据库
Dim BUYdate As String   '购水日期
  BUYdate = Format(CDate(Now), "yyyy-MM-dd HH:mm:ss")
Dim BUYid As String     '购水编号

Set rst = mconn.Execute("select count(yb_id) from wtbddb")
If rst.Fields(0) = 0 Then
    BUYid = "0000001"
Else
    Set rst1 = mconn.Execute("select max(yb_buyid) from WTBDdb")
        If rst1.BOF Then
        Else
        BUYid = FormatString((Val(rst1.Fields(0)) + 1), 7)
        End If
    rst1.Close
End If
rst.Close
Dim BUYnum As String
  BUYnum = "000001"
  
  mconn.Execute ("insert into WTBDdb(yb_buyid,yb_id,yb_type,yb_dn,yb_tdn,yb_num,yb_money,yb_oper,yb_date) values ('" + BUYid + "'," _
                & "'" + Trim(Text15) + "','" + Combo1 + "','" + Text1 + "','" + Text1 + "','" + BUYnum + "','" + Text8 + "'," _
                & "'" + gUserno + "','" + BUYdate + "')")


MsgBox "开户成功！"
  Unload Me
ExitIC

Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim Tempno As String            '用户编号
If Option1.Value Then           '按用户编号
    If Text15 = "" Then
    MsgBox "请输入用户编号"
    Exit Sub
    End If
    
    If Len(Text15) < 7 Then
    Text15 = FormatString(Text15, 7)
    End If

Tempno = Text15
ElseIf Option2.Value Then       '按身份证号码
    If Text10 = "" Then
    MsgBox "请输入用户身份证号码"
    Exit Sub
    End If
Set rst = mconn.Execute("select y_no from YHdb where y_id='" + Text10 + "'")
    If rst.EOF Then
    MsgBox "没有这个身份证号，请检查！"
    Exit Sub
    Else
    Tempno = rst.Fields(0)
    End If
    rst.Close
End If

'由用户编号获取用户详细信息
Set rst = mconn.Execute("select * from YHdb where y_no='" + Tempno + "'")
If rst.EOF Then
    MsgBox "没有此用户的信息，请确认是否输错，或重新添加此用户信息。"
    Frame2.Enabled = False
    Command1.Enabled = False
    Text15.SetFocus
    Exit Sub
Else
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "小区" & Trim(rst.Fields("y_dong")) & "幢" & Trim(rst.Fields("y_dy")) & "单元" & Trim(rst.Fields("y_hao")) & "号"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    Text10 = rst.Fields("y_id")
    Frame2.Enabled = True
    Command1.Enabled = True
End If
rst.Close
Text6 = Format(CDate(Now), "yyyy-MM-dd HH:mm:ss")
'判断是否开过户
Set rst = mconn.Execute("select 1 from WTBDdb where yb_id='" + Tempno + "'")
If Not rst.EOF Then
    MsgBox "此用户已经开过户！不能重复开户"
    Frame2.Enabled = False
    Command1.Enabled = False
    rst.Close
    Exit Sub
End If
rst.Close
Command1.Enabled = True
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'填充用电类型
Set rst = mconn.Execute("select Ds_name from WTDdb")
If rst.EOF Then
MsgBox ("还没有设置用电类型，请先设置好再开户！")
Exit Sub
Else
  Do While Not rst.EOF
    Combo1.AddItem rst.Fields(0)
    rst.MoveNext
  Loop
rst.Close
End If
'获取开户费
Set rst = mconn.Execute("select khfee from SYSdate")
Text16 = rst.Fields(0)
rst.Close
Command1.Enabled = False
End Sub

Private Sub Option1_Click()
Text15.SetFocus
Text10.Locked = True
Text15.Locked = False
End Sub
Private Sub Option2_Click()
Text10.SetFocus
Text15.Locked = True
Text10.Locked = False
End Sub

Private Sub Text1_Change()
On Error GoTo ErR
Text8 = Format(Text1 * Text2, "###.##")
Exit Sub
ErR:
MsgBox ("请先选择合适的用电类型！")
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 27 Then   'ESC键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If

End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    Call Command3_Click
    Exit Sub
End If

End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9)
End Sub


Private Sub Text9_Change()
If Text9 = "" Then
Exit Sub
Else
SkinLabel17.Caption = "找零" & Str(Val(Text9) - Val(Text8))
End If
End Sub
Private Sub Text9_LostFocus()
If Text9 = "" Then
Text9 = "0"
End If
End Sub

