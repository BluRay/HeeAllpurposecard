VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmYhKh 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "FrmYhKh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10500
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ˮ��Ϣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   10215
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   8400
         TabIndex        =   62
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   61
         Text            =   "0"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   7200
         TabIndex        =   60
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   8400
         TabIndex        =   56
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   55
         Text            =   "0"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7200
         TabIndex        =   54
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   50
         Text            =   "0"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   49
         Text            =   "0"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   48
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   17
         Text            =   "0"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   16
         Text            =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   375
         Index           =   0
         Left            =   4440
         OleObjectBlob   =   "FrmYhKh.frx":6F0C2
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   375
         Index           =   0
         Left            =   3000
         OleObjectBlob   =   "FrmYhKh.frx":6F11F
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Index           =   0
         Left            =   1440
         OleObjectBlob   =   "FrmYhKh.frx":6F17C
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   375
         Index           =   0
         Left            =   5880
         OleObjectBlob   =   "FrmYhKh.frx":6F1D9
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Index           =   0
         Left            =   7200
         OleObjectBlob   =   "FrmYhKh.frx":6F234
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   375
         Index           =   0
         Left            =   8400
         OleObjectBlob   =   "FrmYhKh.frx":6F291
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   3
         Left            =   480
         OleObjectBlob   =   "FrmYhKh.frx":6F2EC
         TabIndex        =   66
         Top             =   1560
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   2
         Left            =   480
         OleObjectBlob   =   "FrmYhKh.frx":6F349
         TabIndex        =   65
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   1
         Left            =   480
         OleObjectBlob   =   "FrmYhKh.frx":6F3A6
         TabIndex        =   64
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   0
         Left            =   480
         OleObjectBlob   =   "FrmYhKh.frx":6F403
         TabIndex        =   63
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ѡ��Ҫ������ˮ��ţ�"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   10215
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   7800
         TabIndex        =   6
         ToolTipText     =   "��ˮ����ʾδ���ã��뵽ϵͳ���ò˵��½�������"
         Top             =   220
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4880
         TabIndex        =   5
         ToolTipText     =   "��ˮ����ʾδ���ã��뵽ϵͳ���ò˵��½�������"
         Top             =   220
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2440
         TabIndex        =   4
         ToolTipText     =   "��ˮ����ʾδ���ã��뵽ϵͳ���ò˵��½�������"
         Top             =   220
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ��һ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "��ˮ����ʾδ���ã��뵽ϵͳ���ò˵��½�������"
         Top             =   220
         Width           =   2775
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9960
      OleObjectBlob   =   "FrmYhKh.frx":6F460
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "FrmYhKh.frx":6F694
      TabIndex        =   38
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2880
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   27
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmYhKh.frx":6F6FC
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ������"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ������"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   8
      Text            =   "0"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û���Ϣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   10215
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�����֤���뿪����"
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "��ʾ���û����ǰ��0��ʡ��"
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���û���ſ�����"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ȷ ��"
         Height          =   375
         Left            =   7920
         TabIndex        =   2
         Top             =   160
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   9975
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   400
            Width           =   4215
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   160
            Width           =   4215
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   400
            Width           =   1575
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   160
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   4200
            OleObjectBlob   =   "FrmYhKh.frx":6FFA9
            TabIndex        =   28
            Top             =   405
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   4200
            OleObjectBlob   =   "FrmYhKh.frx":7001D
            TabIndex        =   29
            Top             =   165
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmYhKh.frx":70085
            TabIndex        =   30
            Top             =   405
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmYhKh.frx":700ED
            TabIndex        =   31
            Top             =   165
            Width           =   975
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   4200
      OleObjectBlob   =   "FrmYhKh.frx":70155
      TabIndex        =   26
      Top             =   120
      Width           =   3135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   375
      Left            =   3720
      OleObjectBlob   =   "FrmYhKh.frx":701B6
      TabIndex        =   39
      Top             =   5160
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   375
      Left            =   7320
      OleObjectBlob   =   "FrmYhKh.frx":70215
      TabIndex        =   40
      Top             =   5640
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   375
      Left            =   6960
      OleObjectBlob   =   "FrmYhKh.frx":70272
      TabIndex        =   41
      Top             =   5160
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmYhKh.frx":702D3
      TabIndex        =   42
      Top             =   5640
      Width           =   3495
   End
End
Attribute VB_Name = "FrmYhKh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim rst As Recordset
Dim rst1 As Recordset
Dim buyShu As Integer
Dim oldpass As String * 4
Dim password(1) As Byte
Dim Para(59) As Byte   '�������飬��0�ֽ�

Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value Then
    Text1(Index).Enabled = True
    Text2(Index).Enabled = True
    Text3(Index).Enabled = True
    Text4(Index).Enabled = True
    Text5(Index).Enabled = True
    Text6(Index).Enabled = True
Else
    Text1(Index).Enabled = False
    Text2(Index).Enabled = False
    Text3(Index).Enabled = False
    Text4(Index).Enabled = False
    Text5(Index).Enabled = False
    Text6(Index).Enabled = False
End If
If JTYes Then
For i = 0 To 3
Text4(i).Enabled = False
Next i
End If
End Sub

Private Sub Command1_Click()    '����
On Error GoTo errhandle

'�ܹ�ˮ������Ϊ0
If (Val(Text6(0)) + Val(Text6(1)) + Val(Text6(2)) + Val(Text6(3))) = 0 Then
MsgBox ("����û�й�ˮ")
Exit Sub
End If
'�ж�IC���Ƿ�׼����
If Not InitICcard Then
    ExitIC
    Exit Sub
End If
st = chk_102(icdev)             '�����Ƿ�Ϊ�Ϸ���
If st <> 0 Then
    MsgBox ("���ǺϷ���IC�������顣")
    Exit Sub
End If

'***************�˶�����f0f0***************************
'password(0) = &HF0
'password(1) = &HF0
'Ҫ���������룬��ΪҪ���ǵ������Ƚ��е�����������Ѿ����1b6c
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    oldpass = "1b6c"
    st = asc_hex(oldpass, password(0), 2)
    st = csc_102(icdev, 2, password(0))
    If st < 0 Then
    MsgBox ("�˶�IC����������������¿������忨������������ʹ���𻵣�")
    Exit Sub
    End If
End If
'**************����0��******************************
st = ser_102(icdev, 0, 18, 5)
If st < 0 Then
    MsgBox ("����������")
    Exit Sub
End If
'д��ϵͳ����־
Para(0) = &H98
st = swr_102_hex(icdev, 0, 21, 1, Para(0))
If st < 0 Then
MsgBox "д��ʧ�ܣ�"
Exit Sub
End If
''*************д�û�����־***********************
Para(0) = &H10
st = swr_102_hex(icdev, 0, 18, 1, Para(0))
If st < 0 Then
MsgBox "д��ʧ�ܣ�"
Exit Sub
End If
'**************�����û�������******************************
'����
Set rst = mconn.Execute("select area from Sysdate")
st = asc_hex(rst.Fields(0), Para(0), 2)
rst.Close
'�û����
Para(2) = Val(Text15) / 256
Para(3) = Val(Text15) Mod 256
'��ˮ��Ϣ
If JTYes Then       '�Ƿ�ͨ����ˮ��
    For i = 0 To 3
        If Check1(i).Value Then     '�����ô˱�
        Para(4 + i * 2) = Val(Text6(i)) Mod 256         '��ˮ����� ��λ��ǰ
        Para(5 + i * 2) = Val(Text6(i)) \ 256           '��ˮ����� ��λ�ں�
        Else
        Para(4 + i * 2) = &H0                       '
        Para(5 + i * 2) = &H0                       '
        End If
    Next i
Else
    For i = 0 To 3
        If Check1(i).Value Then     '�����ô˱�
        buyShu = Val(Text4(i)) * 100    '��ˮ������0.01��Ϊ��λ---����ͨ����ˮ�ۣ���ͨ����ǮΪ��λ
        Para(4 + i * 2) = buyShu \ 100 Mod 256     '��ˮ����������
        Para(5 + i * 2) = buyShu Mod 100           '��ˮ��С������
        Else
        Para(4 + i * 2) = &H0                       '
        Para(5 + i * 2) = &H0                       '
        End If
    Next i
End If
'**************����*************************
For i = 0 To 3
    If Check1(i).Value Then     '�����ô˱�
    Para(12 + i) = &H1       '����
    Else
    Para(12 + i) = &H0                       '����
    End If
Next i
'***16-20����ʱ�����****************
For i = 16 To 20
Para(i) = &H0
Next i
'��ˮ��λ21,�����ͨ����ˮ�ۣ����ֽ���00
If JTYes Then
    Para(21) = &H0
Else
    Dim Btemp2 As Integer
    Btemp2 = 0
    For i = 0 To 3
        If Val(Text4(i)) > 255 Then
        Btemp2 = Btemp2 + 2 ^ i
        End If
    Next i
    Para(21) = Btemp2
End If
'������־22
Para(22) = &H11
'����ʱ���־23
Para(23) = &H0
'������־24
Para(24) = &H0
'�հ�25-----��ͨ���ݱ�־
If JTYes Then
Para(25) = &H11
Else
Para(25) = &H0
End If
'��Ź�ˮ��־26     ----ex:12��Ϊ0011
Dim Btemp As Integer
Btemp = 0
For i = 0 To 3
    If Check1(i).Value Then
    Btemp = Btemp + 2 ^ i
    End If
Next i
Para(26) = Btemp
'У��27
Para(27) = &H0
For i = 0 To 26
    Para(27) = Para(27) Xor Para(i)
Next i

'����λ������������
For i = 28 To 59
    Para(i) = &HFF
Next i


'**************����������ϣ���ʼд��****************
'**************������ַ******************************
st = ser_102(icdev, 2, 0, 60)
If st < 0 Then
    MsgBox ("����ʧ�ܣ�")
    Exit Sub
End If

Screen.MousePointer = vbHourglass

st = swr_102_hex(icdev, 2, 2, 20, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
st = swr_102_hex(icdev, 2, 22, 40, Para(20))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
  Screen.MousePointer = vbDefault
  Exit Sub
End If

  Screen.MousePointer = vbDefault
  
'*********���濪�����ݵ����ݿ���******************
 Dim WTopen As String   '����ʱ���ñ���
 WTopen = ""
 For i = 3 To 0 Step -1
    If Check1(i).Value Then
    WTopen = WTopen & "1"
    Else
    WTopen = WTopen & "0"
    End If
 Next i
Dim BUYdate As String   '��ˮ����
  BUYdate = Format(CDate(Now), "yyyy-MM-dd HH:mm:ss")
Dim BUYid As String     '��ˮ��ŵ�һ��
Set rst = mconn.Execute("select count(yb_id) from wtbdb")
    If rst.Fields(0) = 0 Then '�״ο���
    BUYid = "0000001"
    Else
    Set rst1 = mconn.Execute("select max(yb_buyid) from WTBdb")
        If Not rst1.BOF Then
        BUYid = FormatString((Val(rst1.Fields(0)) + 1), 7)
        End If
    rst1.Close
    End If
rst.Close

Dim BUYnum As String
  BUYnum = "000001"

  mconn.Execute ("insert into WTBdb(yb_buyid,yb_id,yb_open,yb_w1,yb_w2,yb_w3,yb_w4,yb_tw1,yb_tw2,yb_tw3,yb_tw4,yb_wdi1,yb_wdi2,yb_wdi3,yb_wdi4,yb_num,yb_money,yb_operator,yb_date) values ('" + BUYid + "'," _
                & "'" + Trim(Text15) + "','" + WTopen + "','" + Text4(0) + "','" + Text4(1) + "','" + Text4(2) + "','" + Text4(3) + "'," _
                & "'" + Text4(0) + "','" + Text4(1) + "','" + Text4(2) + "','" + Text4(3) + "','" + Text5(0) + "','" + Text5(1) + "','" + Text5(2) + "','" + Text5(3) + "'," _
                & "'" + BUYnum + "','" + Text8 + "','" + gUserno + "','" + BUYdate + "')")
'***************************************************************************************************************************************************************************************************************************************************************00
'*************��������************************
'2��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����2���ж�����*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("������λ�������")
  Exit Sub
End If
'***************************************************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("���Ŀ��������")
    Exit Sub
End If

'*********���û�����д�������***************************
Dim Apass As String
Set rst = mconn.Execute("select Apass from Sysdate")
Apass = rst.Fields(0)
'Call ToBCD(Val(Left(Apass, 2)), password(0))
'Call ToBCD(Val(Right(Apass, 2)), password(1))
st = asc_hex(Apass, password(0), 2)
If st < 0 Then
    MsgBox ("��ȡ�������")
    Exit Sub
End If
rst.Close

st = ser_102(icdev, 2, 84, 2)
If st < 0 Then
    MsgBox ("����ʧ�ܣ�")
    Exit Sub
End If
st = swr_102_hex(icdev, 2, 84, 2, password(0))
If st < 0 Then
    MsgBox ("д�������������")
    Exit Sub
End If
'******************************************************

MsgBox "�����ɹ���"
  Unload Me
ExitIC
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()    '�������֤����õ��û���Ϣ
On Error GoTo errhandle
Dim Tempid As String
If Option1.Value Then           '���û����
If Len(Text15) < 7 Then
Text15 = FormatString(Text15, 7)
End If

Set rst = mconn.Execute("select y_id from YHdb where y_no='" + Text15 + "'")
If rst.EOF Then
    MsgBox "û�������ŵ���Ϣ����ȷ���Ƿ������������Ӵ��û���Ϣ��"
    Text15.SetFocus
    Exit Sub
End If
Tempid = rst.Fields("y_id")
rst.Close
ElseIf Option2.Value Then
Tempid = Text10
End If


Set rst = mconn.Execute("select * from YHdb where y_id='" + Tempid + "'")
If rst.EOF Then
    MsgBox "û��������֤�������Ϣ����ȷ���Ƿ������������Ӵ��û���Ϣ��"
    Frame3.Enabled = False
    Command1.Enabled = False
    Text10.SetFocus
    Exit Sub
Else
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "С��" & Trim(rst.Fields("y_dong")) & "��" & Trim(rst.Fields("y_dy")) & "��Ԫ" & Trim(rst.Fields("y_hao")) & "��"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    Frame3.Enabled = True
    Command1.Enabled = True
End If

rst.Close
'�ж��Ƿ��Ѿ�����
Set rst = mconn.Execute("select 1 from WTBdb where yb_id='" + Trim(Text15) + "'")
If Not rst.EOF Then
MsgBox "���û��Ѿ��������������ظ�����"
Frame3.Enabled = False
Command1.Enabled = False

Exit Sub
End If
rst.Close
'********************
'�ж���Щ��û�����ò��������ܹ�ˮ
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='1'")
If rst.Fields(0) = "          " Then
Check1(0).Caption = "ˮ��һ(δ����)"
Check1(0).Enabled = False
Else
Check1(0).Caption = "ˮ��һ��" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(0) = rst1.Fields("w_price")
    Text2(0) = rst.Fields(0)
    Text3(0) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='2'")
If rst.Fields(0) = "          " Then
Check1(1).Caption = "ˮ���(δ����)"
Check1(1).Enabled = False
Else
Check1(1).Caption = "ˮ�����" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(1) = rst1.Fields("w_price")
    Text2(1) = rst.Fields(0)
    Text3(1) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='3'")
If rst.Fields(0) = "          " Then
Check1(2).Caption = "ˮ����(δ����)"
Check1(2).Enabled = False
Else
Check1(2).Caption = "ˮ������" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(2) = rst1.Fields("w_price")
    Text2(2) = rst.Fields(0)
    Text3(2) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='4'")
If rst.Fields(0) = "          " Then
Check1(3).Caption = "ˮ����(δ����)"
Check1(3).Enabled = False
Else
Check1(3).Caption = "ˮ���ģ�" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(3) = rst1.Fields("w_price")
    Text2(3) = rst.Fields(0)
    Text3(3) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Check1(0).Refresh
Check1(1).Refresh
Check1(2).Refresh
Check1(3).Refresh

'*****************************************************************************
'***************�Ƿ�ͨ����ˮ��,��ͨ��Ǯ��ˮ*******************************
If JTYes Then
Frame2(0).Caption = "��ˮ��Ϣ��    ��ǰ�ѿ�ͨ����ˮ�ۣ��밴��ˮ��"
For i = 0 To 3          '��ʾ����ˮ������ͼ�
Set rst = mconn.Execute("select jia1 from Sysjt")
Text1(i) = Val(rst.Fields(0))
Next i
End If


Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

'�������úò���
Set rst = mconn.Execute("select count(wt_type) from wtsdb ")
If rst.Fields(0) = 0 Then
MsgBox "�������ò���"
Command1.Enabled = False
End If
rst.Close

'��ȡ������
Set rst = mconn.Execute("select khfee from SYSdate")
Text7 = rst.Fields(0)
rst.Close
End Sub

Private Sub Option1_Click()
Text15.SetFocus
End Sub
Private Sub Option2_Click()
Text10.SetFocus
End Sub

'Private Sub Text4_Click(Index As Integer)
'Text4(Index) = ""
'End Sub

Private Sub Text4_GotFocus(Index As Integer)
Text4(Index).SelStart = 0
Text4(Index).SelLength = Len(Text4(Index))
End Sub
Private Sub Text6_GotFocus(Index As Integer)
Text6(Index).SelStart = 0
Text6(Index).SelLength = Len(Text6(Index))
End Sub

Private Sub Text5_Click(Index As Integer)
Text4(Index) = ""
End Sub
Private Sub Text6_LostFocus(Index As Integer)
On Error GoTo ErrH
If Text6(Index) = "" Then
    Text6(Index) = "0"
    Exit Sub
ElseIf Val(Text6(Index)) = 0 Then
    Text6(Index) = "0"
    Text4(Index) = "0"
Else
    Text4(Index) = Format((Text6(Index) / Text1(Index)), "####.#")
    Text8 = Format((Val(Text7) + Val(Text6(0)) + Val(Text6(1)) + Val(Text6(2)) + Val(Text6(3))), "#####.#")
End If
Exit Sub
ErrH:
Text6(Index) = "0"
Text4(Index) = "0"
End Sub
Private Sub Text4_LostFocus(Index As Integer)
On Error GoTo ErrH
If Text4(Index) = "" Then
    Text4(Index) = "0"
    Exit Sub
ElseIf Val(Text4(Index)) = 0 Then
    Text6(Index) = "0"
    Text4(Index) = "0"
Else
    Text6(Index) = Format((Text4(Index) * Text1(Index)), "####.#")
    Text8 = Format((Val(Text7) + Val(Text6(0)) + Val(Text6(1)) + Val(Text6(2)) + Val(Text6(3))), "#####.#")
End If
Exit Sub
ErrH:
Text4(Index) = "0"
Text6(Index) = "0"
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 27 Then   'ESC��
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If

End Sub


Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 27 Then   'ESC��
    KeyAscii = 0
    SendKeys "{tab}"
    Exit Sub
End If

End Sub
Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 27 Then   'ESC��
    KeyAscii = 0
    SendKeys "{tab}"
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
SkinLabel17.Caption = "����" & Str(Val(Text9) - Val(Text8))
End If
End Sub
Private Sub Text9_LostFocus()
If Text9 = "" Then
Text9 = "0"
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
KeyAscii = limitnumber(KeyAscii)  'ֻ��Ϊ����
 If KeyAscii = 13 Then   '�س���
    Call Command3_Click
    Exit Sub
End If

End Sub

