VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmJieTi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "阶梯水价设置"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8700
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "修改水价"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2400
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   13
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmJieTi.frx":0000
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "当前未开通阶梯水价！"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   8295
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":05D3
         TabIndex        =   9
         Top             =   1440
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":0711
         TabIndex        =   10
         Top             =   960
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "FrmJieTi.frx":084F
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启用阶梯水价"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3480
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmJieTi.frx":08F5
      Top             =   3720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   495
      Left            =   3480
      OleObjectBlob   =   "FrmJieTi.frx":0B29
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmJieTi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset

Private Sub Command1_Click()
On Error GoTo errhandle
'**********警告信息*******************
If MsgBox("修改阶梯水价将影响之后所有购水操作！且设置后需对所有水表进行重新设置！" + Chr(13) + "您确认要修改吗？", vbYesNo, "警告！！！") = vbNo Then
Exit Sub
End If

'**********启用/停用阶梯水价**********
If JTYes Then   '如果当前状态为开启阶梯水价，则此处为关闭
    mconn.Execute ("update sysjt set jietiyesno='no'")
    JTYes = False
Else            '开启阶梯水价
'输入的阶梯水价是否合法
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
        MsgBox "所有项目必须填写完整！！"
        Exit Sub
    End If
    If Val(Text3) <> Val(Text5) Then
    MsgBox "您填写的数据有误，请检查！"
    Text3.SetFocus
    Exit Sub
    End If
mconn.Execute ("update sysjt set jietiyesno='yes',jia1='" + Text1 + "',jia2='" + Text4 + "',jia3='" + Text6 + "',nian1='" + Text2 + "',nian2='" + Text3 + "'")
JTYes = True
End If
MsgBox "设置成功！！！"
Unload Me
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo errhandle
'**********警告信息*******************
If MsgBox("修改阶梯水价将影响之后所有购水操作！且设置后需对所有水表进行重新设置！" + Chr(13) + "您确认要修改吗？", vbYesNo, "警告！！！") = vbNo Then
Exit Sub
End If

'输入的阶梯水价是否合法
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
        MsgBox "所有项目必须填写完整！！"
        Exit Sub
    End If
    If Val(Text3) <> Val(Text5) Then
    MsgBox "您填写的数据有误，请检查！"
    Text3.SetFocus
    Exit Sub
    End If
mconn.Execute ("update sysjt set jietiyesno='yes',jia1='" + Text1 + "',jia2='" + Text4 + "',jia3='" + Text6 + "',nian1='" + Text2 + "',nian2='" + Text3 + "'")
JTYes = True
MsgBox "修改成功！！"



Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'**********当前是否已经开通阶梯水价**********
If JTYes Then
    Frame1.Caption = "当前已经开通阶梯水价！"
    Command1.Caption = "停用阶梯水价"
    Command3.Visible = True
    '填充当前阶梯水价数据
    Set rst = mconn.Execute("select * from Sysjt")
    Text1 = rst.Fields("jia1")
    Text2 = rst.Fields("nian1")
    Text3 = rst.Fields("nian2")
    Text4 = rst.Fields("jia2")
    Text5 = rst.Fields("nian2")
    Text6 = rst.Fields("jia3")
    rst.Close
    
Else
    JTYes = False
    Frame1.Caption = "当前未开通阶梯水价！"
    Command1.Caption = "启用阶梯水价"
    Command3.Visible = False
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   '小数点
    KeyAscii = 0
    MsgBox "只能为整数！"
    Exit Sub
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   '小数点
    KeyAscii = 0
    MsgBox "只能为整数！"
    Exit Sub
End If
End Sub

Private Sub Text3_LostFocus()
Text5 = Text3
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
 If KeyAscii = 46 Then   '小数点
    KeyAscii = 0
    MsgBox "只能为整数！"
    Exit Sub
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  '只能为数字
 If KeyAscii = 13 Then   '回车键
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If
End Sub

