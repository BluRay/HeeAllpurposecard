VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmRegsys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统注册"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6225
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "FrmRegsys.frx":0000
      Top             =   2520
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "FrmRegsys.frx":0234
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "FrmRegsys.frx":029A
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   1680
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmRegsys.frx":0304
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "请向系统提供商索取注册码"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   495
      Left            =   2520
      OleObjectBlob   =   "FrmRegsys.frx":0CB8
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmRegsys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Skey = power(Miwen, Smy)
If Trim(Text2) = Skey Then
MsgBox "系统注册成功！请重新启动本系统！"
'更新数据库
'mconn.Execute ("delete from regsys")
mconn.Execute ("insert into regsys (HDnum,Skey,shiyongcishu) values('" + Miwen + "','" + Skey + "','0')")
Unload Me

Else
MsgBox "系统注册失败！"
Unload Me
Call QuitSystem
End If
End Sub

Private Sub Command2_Click()
Unload Me
Call QuitSystem
End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
    Text1 = Miwen
'    Text2.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Call QuitSystem
End Sub
