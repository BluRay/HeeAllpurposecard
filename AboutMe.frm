VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form AboutMe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ǰ����Ա��Ϣ"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6735
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ  ��"
      Height          =   615
      Left            =   2460
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Ӫҵ����Ȩ"
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1995
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " �����޸�Ȩ"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ���ݱ��ݺͻָ�Ȩ"
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   2385
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ����IC������Ȩ"
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   2385
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   1920
      OleObjectBlob   =   "AboutMe.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "AboutMe.frx":005F
      Top             =   3120
   End
End
Attribute VB_Name = "AboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim TempPower As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
SkinLabel1.Caption = gUsername & "���е�Ȩ��"
'��ȡȨ��
Set rst = mconn.Execute("select power from operator where operatorno='" + gUserno + "'")
TempPower = rst.Fields(0)
        If InStr(TempPower, "A") <> 0 Then    'ӵ�г�Ʊ����Ȩ
            ChkOppower(0).Value = 1
        Else
            ChkOppower(0).Value = 0
        End If
        If InStr(TempPower, "B") <> 0 Then    'ӵ�в������Ȩ
            ChkOppower(1).Value = 1
        Else
            ChkOppower(1).Value = 0
        End If
        If InStr(TempPower, "C") <> 0 Then     'ӵ�����ݱ��ݺͻָ�Ȩ
            ChkOppower(2).Value = 1
        Else
            ChkOppower(2).Value = 0
        End If
        If InStr(TempPower, "D") <> 0 Then     'IC���ع�Ȩ
            ChkOppower(3).Value = 1
        Else
            ChkOppower(3).Value = 0
        End If

End Sub
