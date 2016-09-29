VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmRdCd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "读取IC卡信息"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "FrmRdCd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9285
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8520
      OleObjectBlob   =   "FrmRdCd.frx":030A
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2880
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   17
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmRdCd.frx":053E
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "设置卡参数："
      Height          =   2175
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "FrmRdCd.frx":0DEB
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmRdCd.frx":0E55
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmRdCd.frx":0EBF
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "FrmRdCd.frx":0F27
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7095
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   4200
            OleObjectBlob   =   "FrmRdCd.frx":0F8F
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   720
            OleObjectBlob   =   "FrmRdCd.frx":0FF5
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "FrmRdCd.frx":105B
            Left            =   1680
            List            =   "FrmRdCd.frx":1071
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "FrmRdCd.frx":1099
            Left            =   5040
            List            =   "FrmRdCd.frx":10A6
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9015
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   6840
         OleObjectBlob   =   "FrmRdCd.frx":10BA
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "FrmRdCd.frx":112A
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmRdCd.frx":1196
         TabIndex        =   26
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmRdCd.frx":11FE
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   4560
         TabIndex        =   5
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   4560
         TabIndex        =   4
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmRdCd.frx":1266
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "FrmRdCd.frx":12D0
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "FrmRdCd.frx":1338
         TabIndex        =   28
         Top             =   840
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   3720
      OleObjectBlob   =   "FrmRdCd.frx":13A0
      TabIndex        =   16
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmRdCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldpass As String * 4
Dim password(1) As Byte
Dim RdType(9) As Byte, YHid(1) As Byte
Dim rst As Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errhandle
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
Dim i As Integer
'判断IC卡是否准备好
If Not InitICcard Then
    Exit Sub
End If
st = chk_102(icdev)             '测试是否为合法卡
If st <> 0 Then
    MsgBox ("读IC卡错误！！请检查。")
    Frame4.Visible = False
    Frame1.Visible = True
    Exit Sub
End If
'***************核对密码f0f0***************************
oldpass = "1b6c"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("核对IC卡密码错")
    Exit Sub
End If

Dim idTemp As String

'读卡信息????
st = srd_102_hex(icdev, 0, 18, 1, RdType(0))
If RdType(0) = &H10 Then
    MsgBox "此卡为用户卡"
    Frame4.Visible = True
    Frame1.Visible = False
    '填充用户详细信息
    '读取用户编号
    st = srd_102_hex(icdev, 2, 4, 2, YHid(0))
    idTemp = FormatString(Str(YHid(0) * 256 + YHid(1)), 7)
    '填充用户信息
    Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
    Text10 = rst.Fields("y_id")
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "小区" & Trim(rst.Fields("y_dong")) & "幢" & Trim(rst.Fields("y_dy")) & "单元" & Trim(rst.Fields("y_hao")) & "号"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    rst.Close
    Set rst = mconn.Execute("select count(*)from WTBDB where yb_id='" + idTemp + "'")
    SkinLabel14.Caption = "该用户已经购买" & Str(rst.Fields(0)) & "次水"
    rst.Close
ElseIf RdType(0) = &H50 Then
    MsgBox "此卡为初始化卡！"
    Frame4.Visible = True
    Frame1.Visible = False
ElseIf RdType(0) = &H30 Then
    MsgBox "此卡为清零卡！"
    Frame4.Visible = True
    Frame1.Visible = False
ElseIf RdType(0) = &H40 Then
    MsgBox "此卡为查询卡！"
    Frame4.Visible = True
    Frame1.Visible = False
ElseIf RdType(0) = &H20 Then
    MsgBox "此卡为设置卡！"
    Frame4.Visible = False
    Frame1.Visible = True
    '填充设置卡信息
    Dim setCd(8) As Byte
    st = srd_102_hex(icdev, 2, 1, 9, RdType(0))
    Select Case RdType(2)
        Case 1
        Combo1 = "表一"
        Case 2
        Combo1 = "表二"
        Case 3
        Combo1 = "表三"
        Case 4
        Combo1 = "表四"
        Case 5
        Combo1 = "表五"
        Case 6
        Combo1 = "表六"
    End Select
    Select Case RdType(3)
        Case 1
        Combo2 = "1T"
        Case 2
        Combo2 = "0.1T"
        Case 3
        Combo2 = "0.5T"
    End Select
    Text1 = RdType(8)
    Text2 = RdType(4) * 256 + RdType(5)
    Text3 = RdType(6)
    Text4 = RdType(7)
Else
'    MsgBox "此卡类型无法识别！"
'    Frame4.Visible = False
'    Frame1.Visible = True
'读电表卡
    st = srd_102_hex(icdev, 0, 2, 5, setCd(0))
    If st < 0 Then
    MsgBox "读卡失败！"
    Exit Sub
    End If
    If setCd(4) = &H41 Then
    MsgBox "此卡为电表卡"
    Call BCDTo(setCd(2), idTemp)
    Dim idTemp2 As String
    Call BCDTo(setCd(3), idTemp2)
    idTemp = FormatString(Val(idTemp & idTemp2), 7)
    End If
    Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
    Text10 = rst.Fields("y_id")
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "小区" & Trim(rst.Fields("y_dong")) & "幢" & Trim(rst.Fields("y_dy")) & "单元" & Trim(rst.Fields("y_hao")) & "号"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    rst.Close
    Set rst = mconn.Execute("select count(*)from WTBDdB where yb_id='" + idTemp + "'")
    SkinLabel14.Caption = "该用户已经购买" & Str(rst.Fields(0)) & "次电"
    rst.Close

    
End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Exit Sub
End Sub


