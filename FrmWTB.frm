VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmWTB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û��ճ���ˮ"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "FrmWTB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10470
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5640
      TabIndex        =   0
      Text            =   "0"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ˮ��Ϣ��"
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
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   2400
      Width           =   10215
      Begin VB.TextBox Text1 
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
         Height          =   375
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Index           =   0
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   6960
         TabIndex        =   47
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   46
         Text            =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   8160
         TabIndex        =   45
         Text            =   "0"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Index           =   1
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   6960
         TabIndex        =   41
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   40
         Text            =   "0"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   8160
         TabIndex        =   39
         Text            =   "0"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Index           =   2
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   6960
         TabIndex        =   35
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   34
         Text            =   "0"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   8160
         TabIndex        =   33
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
         Height          =   375
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Index           =   3
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   6960
         TabIndex        =   29
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   28
         Text            =   "0"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   8160
         TabIndex        =   27
         Text            =   "0"
         Top             =   1560
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   375
         Index           =   1
         Left            =   4200
         OleObjectBlob   =   "FrmWTB.frx":030A
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   375
         Index           =   1
         Left            =   2760
         OleObjectBlob   =   "FrmWTB.frx":0367
         TabIndex        =   52
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Index           =   1
         Left            =   1200
         OleObjectBlob   =   "FrmWTB.frx":03C4
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   375
         Index           =   1
         Left            =   5640
         OleObjectBlob   =   "FrmWTB.frx":0421
         TabIndex        =   54
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Index           =   1
         Left            =   6960
         OleObjectBlob   =   "FrmWTB.frx":047C
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   375
         Index           =   1
         Left            =   8160
         OleObjectBlob   =   "FrmWTB.frx":04D9
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   3
         Left            =   240
         OleObjectBlob   =   "FrmWTB.frx":0534
         TabIndex        =   57
         Top             =   1560
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   2
         Left            =   240
         OleObjectBlob   =   "FrmWTB.frx":0591
         TabIndex        =   58
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   1
         Left            =   240
         OleObjectBlob   =   "FrmWTB.frx":05EE
         TabIndex        =   59
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   4
         Left            =   240
         OleObjectBlob   =   "FrmWTB.frx":064B
         TabIndex        =   60
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2880
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   25
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmWTB.frx":06A8
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9360
      OleObjectBlob   =   "FrmWTB.frx":0F55
      Top             =   5640
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   375
      Left            =   7800
      OleObjectBlob   =   "FrmWTB.frx":1189
      TabIndex        =   23
      Top             =   5160
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   375
      Left            =   7320
      OleObjectBlob   =   "FrmWTB.frx":11E6
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   375
      Left            =   4200
      OleObjectBlob   =   "FrmWTB.frx":1247
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmWTB.frx":12A6
      TabIndex        =   20
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   330
      Left            =   9000
      TabIndex        =   12
      Text            =   "0"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ������"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ������"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û���Ϣ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   10215
      Begin VB.CommandButton Command3 
         Caption         =   "������û�������˶�ȡ�û���Ϣ"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   9855
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   720
            Width           =   4335
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "FrmWTB.frx":1362
            TabIndex        =   13
            Top             =   720
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "FrmWTB.frx":13D6
            TabIndex        =   14
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmWTB.frx":143E
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmWTB.frx":14A6
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "FrmWTB.frx":150E
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "FrmWTB.frx":1576
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   7320
            OleObjectBlob   =   "FrmWTB.frx":15E0
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   3960
      OleObjectBlob   =   "FrmWTB.frx":1650
      TabIndex        =   24
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmWTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldpass As String * 4
Dim password(1) As Byte
Dim RdType As Byte
Dim YHid(3) As Byte
Dim rst As Recordset
Dim rst1 As Recordset
Dim buyShu As Integer
Dim BUYcushu As Integer '��ˮ����
Dim Para(55) As Byte   '�������飬��56�ֽ�
Dim i As Integer

Private Sub Command1_Click()
On Error GoTo errhandle
If Val(Text8) = 0 Then
MsgBox "�㻹û�й�ˮ��"
Exit Sub
End If
'ǰ���ֽ����ݲ���

'**********��ˮ��Ϣ**********
If JTYes Then                         '�Ƿ����ý���ˮ��
    For i = 0 To 3
        If Text6(i).Enabled Then      '�����ô˱�
        Para(0 + i * 2) = Val(Text6(i)) Mod 256         '��ˮ����� ��λ��ǰ
        Para(1 + i * 2) = Val(Text6(i)) \ 256           '��ˮ����� ��λ�ں�
        Else
        Para(0 + i * 2) = &H0                       '
        Para(1 + i * 2) = &H0                       '
        End If
    Next i
Else
    For i = 0 To 3
        If Text6(i).Enabled Then                   '�����ô˱�
        buyShu = Val(Text4(i)) * 100               '��ˮ������0.01��Ϊ��λ---����ͨ����ˮ�ۣ���ͨ����ǮΪ��λ
        Para(0 + i * 2) = buyShu \ 100 Mod 256     '��ˮ����������
        Para(1 + i * 2) = buyShu Mod 100           '��ˮ��С������
        Else
        Para(0 + i * 2) = &H0                      '��ˮ����������
        Para(1 + i * 2) = &H0                      '��ˮ��С������
        End If
    Next i
End If
'**********����******************************
For i = 0 To 3
    If Text6(i).Enabled Then                   '�����ô˱�
    Para(8 + i) = BUYcushu                     '����
    Else
    Para(8 + i) = &H0                          '����
    End If
Next i
'**********ʱ������λ12������ʱ��FF**********
For i = 12 To 26
Para(i) = &H0
Next i
'**********��ˮ��λ17,�����ͨ����ˮ�ۣ����ֽ���00**********
If JTYes Then
    Para(17) = &H0
Else
    Dim Btemp2 As Integer
    Btemp2 = 0
    For i = 0 To 3
        If Val(Text4(i)) > 255 Then
        Btemp2 = Btemp2 + 2 ^ i
        End If
    Next i
    Para(17) = Btemp2
End If
'**********������־18**********
Para(18) = &H0
'**********����ʱ���־19******
Para(19) = &H0
'**********������־20**********
Para(20) = &H0
'**********�հ�25-----��ͨ���ݱ�־
If JTYes Then
Para(21) = &H11
Else
Para(21) = &H0
End If
'**********��Ź�ˮ��־22**********124��Ϊ1101
Dim Btemp As Integer
Btemp = 0
For i = 0 To 3
    If Text6(i).Enabled Then
    Btemp = Btemp + 2 ^ i
    End If
Next i
Para(22) = Btemp
'**********У��********************
Para(23) = &H0
For i = 0 To 3
Para(23) = Para(23) Xor YHid(i)
Next i
For i = 0 To 22
Para(23) = Para(23) Xor Para(i)
Next i
'����λ������������
For i = 24 To 55
    Para(i) = &HFF
Next i




'**************������ַ******************************
st = ser_102(icdev, 2, 0, 64)
If st < 0 Then
    MsgBox ("����ʧ�ܣ�")
    Exit Sub
End If
Screen.MousePointer = vbHourglass
'дǰ��λYHid(0)
st = swr_102_hex(icdev, 2, 2, 4, YHid(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
st = swr_102_hex(icdev, 2, 6, 20, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
st = swr_102_hex(icdev, 2, 26, 36, Para(20))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
  Screen.MousePointer = vbDefault
  Exit Sub
End If
Screen.MousePointer = vbDefault
'2��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����2���ж�����*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("������λ�������")
  Exit Sub
End If

'���濪�����ݵ����ݿ���
 Dim WTopen As String   '����ʱ���ñ���
 WTopen = ""
 For i = 3 To 0 Step -1
    If Text5(i).Enabled Then
    WTopen = WTopen & "1"
    Else
    WTopen = WTopen & "0"
    End If
 Next i
Dim BUYdate As String   '��ˮ����
  BUYdate = Format(CDate(Now), "yyyy-MM-dd HH:mm:ss")
Dim BUYid As String     '��ˮ���
Set rst = mconn.Execute("select max(yb_buyid) from WTBdb")
    If rst.EOF Then
    BUYid = "0000001"
    Else
    BUYid = FormatString((Val(rst.Fields(0)) + 1), 7)
    End If
rst.Close
Dim BUYnum As String
  BUYnum = FormatString(Str(BUYcushu), 6)

  mconn.Execute ("insert into WTBdb(yb_buyid,yb_id,yb_open,yb_w1,yb_w2,yb_w3,yb_w4,yb_tw1,yb_tw2,yb_tw3,yb_tw4,yb_wdi1,yb_wdi2,yb_wdi3,yb_wdi4,yb_num,yb_money,yb_operator,yb_date) values ('" + BUYid + "'," _
                & "'" + Trim(Text15) + "','" + WTopen + "','" + Text4(0) + "','" + Text4(1) + "','" + Text4(2) + "','" + Text4(3) + "'," _
                & "'" + Str(Val(Text5(0)) + Val(Text4(0))) + "','" + Str(Val(Text5(1)) + Val(Text4(1))) + "','" + Str(Val(Text5(2)) + Val(Text4(2))) + "','" + Str(Val(Text5(3)) + Val(Text4(3))) + "','" + Text5(0) + "','" + Text5(1) + "','" + Text5(2) + "','" + Text5(3) + "'," _
                & "'" + BUYnum + "','" + Text8 + "','" + gUserno + "','" + BUYdate + "')")


MsgBox "��ˮ�ɹ���"
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
On Error GoTo errhandle
Dim i As Integer
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

'�˴�Ҫ���αȽ����룬�������Ὣ����ĳ��û����룡����     ????
oldpass = "1b6c"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    Set rst = mconn.Execute("select Apass from Sysdate")
    oldpass = rst.Fields(0)
    rst.Close
    st = asc_hex(oldpass, password(0), 2)
    st = csc_102(icdev, 2, password(0))
    If st < 0 Then
        MsgBox ("�˶�IC�������")
        Exit Sub
    End If
End If
'������Ϣ
st = srd_102_hex(icdev, 0, 18, 1, RdType)
If RdType <> &H10 Then
MsgBox "�˿������û���������ȷ�����û�����"
Exit Sub
End If
'***************��ȡǰ���ֽ�ԭ����******************************
'***************��ȡ�û����******************************
st = srd_102_hex(icdev, 2, 2, 4, YHid(0))
Dim idTemp As String
idTemp = FormatString(Str(YHid(2) * 256 + YHid(3)), 7)


'***************����û���Ϣ******************************
Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
If rst.EOF Then
MsgBox "ϵͳ��û������û�����Ϣ��"
Exit Sub
End If

Text10 = rst.Fields("y_id")
Text11 = rst.Fields("y_name")
Text12 = rst.Fields("y_tel")
Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "С��" & Trim(rst.Fields("y_dong")) & "��" & Trim(rst.Fields("y_dy")) & "��Ԫ" & Trim(rst.Fields("y_hao")) & "��"
Text14 = rst.Fields("y_memo")
Text15 = rst.Fields("y_no")
rst.Close
'***************�ж��û����Ƿ���ˮ����ˢ����ֻ��ˢ�����ܼ���***************
'***************������²������ɿ��Լ���  ???***************
Dim YHwtS(3) As Byte
st = srd_102_hex(icdev, 2, 25, 4, YHwtS(0))      '����ˮ��־
If st < 0 Then
    MsgBox ("����ʧ�ܣ�����")
    Exit Sub
End If
If YHwtS(3) = 0 Then

Else    '������²��Ŀ�����ɼ�������
    If Not YHwtS(1) = &H11 Then
        MsgBox "���û�����δ���ڶ�Ӧ��ˮ����ʹ�ù����޷���������"
        Exit Sub
    End If
End If
'***************��ȡˮ������Ϣ--���ϼ��굥**************  ????






'***************�û���ˮ����******************************
Set rst = mconn.Execute("select count(*)from WTBDB where yb_id='" + idTemp + "'")
BUYcushu = rst.Fields(0) + 1
SkinLabel14.Caption = "���Ǹ��û���" & Str(BUYcushu) & " �ι�ˮ"
rst.Close
'��乺ˮ��Ϣ--�ۼ���
Set rst = mconn.Execute("select * from WTBdb where yb_buyid=(select max(yb_buyid)from WTBdb where yb_id='" + idTemp + "')")

Text5(0) = rst.Fields("yb_tw1")
Text5(1) = rst.Fields("yb_tw2")
Text5(2) = rst.Fields("yb_tw3")
Text5(3) = rst.Fields("yb_tw4")

Dim openTM As String
openTM = rst.Fields("yb_open")
For i = 1 To 4
    If Val(Left(Right(Trim(openTM), i), 1)) Then
    Text1(i - 1).Locked = True
    Text2(i - 1).Locked = True
    Text3(i - 1).Locked = True
    Text4(i - 1).Enabled = True
    Text5(i - 1).Locked = True
    Text6(i - 1).Enabled = True
    Else
    Text1(i - 1).Enabled = False
    Text2(i - 1).Enabled = False
    Text3(i - 1).Enabled = False
    Text4(i - 1).Enabled = False
    Text5(i - 1).Enabled = False
    Text6(i - 1).Enabled = False
    End If
Next i
rst.Close

If JTYes Then
For i = 0 To 3
Text4(i).Enabled = False
Next i
End If

Command1.Enabled = True
Exit Sub
errhandle:
MsgBox (Error(ErR))
Resume Next
End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
'********************
'�������úò���
Set rst = mconn.Execute("select count(wt_type) from wtsdb ")
If rst.Fields(0) = 0 Then
MsgBox "�������ò���"
Exit Sub
End If
rst.Close
'�ж���Щ��û�����ò��������ܹ�ˮ
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='1'")
If rst.Fields(0) = "          " Then
Text1(0).Enabled = False
Text2(0).Enabled = False
Text3(0).Enabled = False
Text4(0).Enabled = False
Text5(0).Enabled = False
Text6(0).Enabled = False
Else
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(0) = rst1.Fields("w_price")
    Text2(0) = rst.Fields(0)
    Text3(0) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='2'")
If rst.Fields(0) = "          " Then
Text1(1).Enabled = False
Text2(1).Enabled = False
Text3(1).Enabled = False
Text4(1).Enabled = False
Text5(1).Enabled = False
Text6(1).Enabled = False
Else
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(1) = rst1.Fields("w_price")
    Text2(1) = rst.Fields(0)
    Text3(1) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='3'")
If rst.Fields(0) = "          " Then
Text1(2).Enabled = False
Text2(2).Enabled = False
Text3(2).Enabled = False
Text4(2).Enabled = False
Text5(2).Enabled = False
Text6(2).Enabled = False
Else
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(2) = rst1.Fields("w_price")
    Text2(2) = rst.Fields(0)
    Text3(2) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='4'")
If rst.Fields(0) = "          " Then
Text1(3).Enabled = False
Text2(3).Enabled = False
Text3(3).Enabled = False
Text4(3).Enabled = False
Text5(3).Enabled = False
Text6(3).Enabled = False
Else
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(3) = rst1.Fields("w_price")
    Text2(3) = rst.Fields(0)
    Text3(3) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
'**********�����ͨ����ˮ�ۣ��򵥼���ʾΪ����ˮ��****************************************
If JTYes Then
    Frame2(1).Caption = "��ˮ��Ϣ��    ��ǰ�ѿ�ͨ����ˮ�ۣ��밴��ˮ��"
    For i = 0 To 3          '��ʾ����ˮ������ͼ�
    Set rst = mconn.Execute("select jia1 from Sysjt")
    Text1(i) = Val(rst.Fields(0))
    Next i
End If


End Sub

Private Sub Text4_LostFocus(Index As Integer)
On Error GoTo ErrH
If Text4(Index) = "" Then
Exit Sub
Else
Text6(Index) = Format(Text4(Index) * Text1(Index), "####.#")
Text8 = Val(Text6(0)) + Val(Text6(1)) + Val(Text6(2)) + Val(Text6(3))
End If
Exit Sub
ErrH:
Text4(Index) = "0"
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
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


Private Sub Text4_GotFocus(Index As Integer)
Text4(Index).SelStart = 0
Text4(Index).SelLength = Len(Text4(Index))
End Sub

Private Sub Text6_GotFocus(Index As Integer)
Text6(Index).SelStart = 0
Text6(Index).SelLength = Len(Text6(Index))
End Sub

Private Sub Text6_LostFocus(Index As Integer)
On Error GoTo ErrH
If Text6(Index) = "" Then
Text6(Index) = "0"
Exit Sub
Else
Text4(Index) = Format((Text6(Index) / Text1(Index)), "####.#")
Text8 = Format((Val(Text9) + Val(Text6(0)) + Val(Text6(1)) + Val(Text6(2)) + Val(Text6(3))), "#####.#")
End If
Exit Sub
ErrH:
Text6(Index) = "0"
End Sub

Private Sub Text9_Change()
If Text9 = "" Then
Exit Sub
Else
SkinLabel17.Caption = "����" & Str(Val(Text9) - Val(Text8))
End If
End Sub

Private Sub Text9_GotFocus()
Text9.SelStart = 0
Text9.SelLength = Len(Text9)
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then
Text9 = "0"
End If
End Sub

