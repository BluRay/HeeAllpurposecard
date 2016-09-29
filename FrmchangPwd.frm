VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmchangPwd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改密码"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "FrmchangPwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5790
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5160
      OleObjectBlob   =   "FrmchangPwd.frx":030A
      Top             =   2760
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmchangPwd.frx":053E
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "FrmchangPwd.frx":05B8
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "FrmchangPwd.frx":0628
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "FrmchangPwd.frx":0698
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   710
      Left            =   1680
      ScaleHeight     =   645
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmchangPwd.frx":070A
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   2520
      OleObjectBlob   =   "FrmchangPwd.frx":0FB7
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   480
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmchangPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text2 = "" Then
MsgBox "密码不能为空！"
Text2.SetFocus
Exit Sub
End If

If Text2 <> Text3 Then
    MsgBox ("您两次输入的新密码不一样，请重新输入！")
    Text2.SetFocus
    Exit Sub
End If
If Text1 = gPassword Then
mconn.Execute ("update operator set password='" + Text3 + "'where operatorno='" + gUserno + "'")
MsgBox ("密码修改成功！")
Unload Me
Else
MsgBox ("您输入的原密码有误！请重新输入")
Text1.SetFocus
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

End Sub
