VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmSorD 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "请选择类别，单击空白处退出"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "水"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "电"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4200
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7290
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmSorD.frx":0000
      Top             =   1800
   End
End
Attribute VB_Name = "FrmSorD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If SorD = "rcgm" Then
FrmWTB.Show vbModal
Unload Me
ElseIf SorD = "yykh" Then
FrmYhKh.Show vbModal
Unload Me
ElseIf SorD = "sysset" Then
MsysSet.Show vbModal
Unload Me
ElseIf SorD = "gmcx" Then
FrmBUYcha.Show vbModal
Unload Me
End If
End Sub

Private Sub Command2_Click()
If SorD = "rcgm" Then
FRMWTBD.Show vbModal
Unload Me
ElseIf SorD = "yykh" Then
FRMYHkhD.Show vbModal
Unload Me
ElseIf SorD = "sysset" Then
MsysSetD.Show vbModal
Unload Me
ElseIf SorD = "gmcx" Then
SorD = "gmcxD"
FrmBUYcha.Show vbModal
Unload Me
End If
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub
