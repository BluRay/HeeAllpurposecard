VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmBuKa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û�����"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "FrmBuKa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9300
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "0"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2280
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   73
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmBuKa.frx":030A
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   6240
      OleObjectBlob   =   "FrmBuKa.frx":0BB7
      TabIndex        =   71
      Top             =   4800
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "FrmBuKa.frx":0C1F
      Top             =   5640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û���Ϣ��"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9015
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   270
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   6720
         OleObjectBlob   =   "FrmBuKa.frx":0E53
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmBuKa.frx":0EBB
         TabIndex        =   41
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ȷ ��"
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   160
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   8175
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "FrmBuKa.frx":0F2F
            TabIndex        =   46
            Top             =   405
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "FrmBuKa.frx":0FA3
            TabIndex        =   45
            Top             =   165
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmBuKa.frx":100B
            TabIndex        =   44
            Top             =   405
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmBuKa.frx":1073
            TabIndex        =   43
            Top             =   165
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   160
            Width           =   1575
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   400
            Width           =   1575
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   160
            Width           =   3375
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   400
            Width           =   3375
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�û����һ�ι�ˮ��Ϣ��"
      Enabled         =   0   'False
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9015
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ�����"
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   1
         Left            =   4800
         TabIndex        =   26
         Top             =   240
         Width           =   3735
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Height          =   270
            Index           =   1
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   1
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   1
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   1
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   1
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":10DB
            TabIndex        =   53
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":1141
            TabIndex        =   54
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":11A7
            TabIndex        =   55
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Index           =   1
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":120D
            TabIndex        =   56
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Index           =   1
            Left            =   1920
            OleObjectBlob   =   "FrmBuKa.frx":1271
            TabIndex        =   57
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Index           =   1
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":12D7
            TabIndex        =   58
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ������"
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   3735
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Height          =   270
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   2
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   2
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   2
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Index           =   2
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":133B
            TabIndex        =   59
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Index           =   2
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":13A1
            TabIndex        =   60
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Index           =   2
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":1407
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Index           =   2
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":146D
            TabIndex        =   62
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Index           =   2
            Left            =   1920
            OleObjectBlob   =   "FrmBuKa.frx":14D1
            TabIndex        =   63
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Index           =   2
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":1537
            TabIndex        =   64
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ���ģ�"
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   3
         Left            =   4800
         TabIndex        =   12
         Top             =   1320
         Width           =   3735
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Height          =   270
            Index           =   3
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   3
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   3
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   3
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   3
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Index           =   3
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":159B
            TabIndex        =   65
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Index           =   3
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":1601
            TabIndex        =   66
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Index           =   3
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":1667
            TabIndex        =   67
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Index           =   3
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":16CD
            TabIndex        =   68
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Index           =   3
            Left            =   1920
            OleObjectBlob   =   "FrmBuKa.frx":1731
            TabIndex        =   69
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Index           =   3
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":1797
            TabIndex        =   70
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ˮ��һ��"
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   3735
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   0
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   0
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":17FB
            TabIndex        =   49
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":1861
            TabIndex        =   48
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "FrmBuKa.frx":18C7
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   0
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Index           =   0
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
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
            Height          =   270
            Index           =   0
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Index           =   0
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":192D
            TabIndex        =   50
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Index           =   0
            Left            =   1920
            OleObjectBlob   =   "FrmBuKa.frx":1991
            TabIndex        =   51
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Index           =   0
            Left            =   2040
            OleObjectBlob   =   "FrmBuKa.frx":19F7
            TabIndex        =   52
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   495
      Left            =   3120
      OleObjectBlob   =   "FrmBuKa.frx":1A5B
      TabIndex        =   72
      Top             =   120
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmBuKa.frx":1ABC
      TabIndex        =   74
      Top             =   4800
      Width           =   1455
   End
End
Attribute VB_Name = "FrmBuKa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset, rst1 As Recordset
Dim i As Integer
Dim oldpass As String * 4
Dim password(1) As Byte
Dim Para(59) As Byte   '�������飬��60�ֽ�
Dim BUYcushu As String, BuyCushuD As String
Dim DisF As Boolean, DisD As Boolean

'?????  ����ʱ�������ݣ���������

Private Sub Command1_Click()
On Error GoTo errhandle
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
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("�˶�IC���������ʹ���¿����Ȼ��տ���")
    Exit Sub
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

If DisF Then    '�Ƿ��й�������
'**************�����û�������******************************
'����0��1
Dim rst As Recordset
Set rst = mconn.Execute("select area from Sysdate")
st = asc_hex(rst.Fields(0), Para(0), 2)
rst.Close
'�û����2��3
Para(2) = Val(Text15) / 256
Para(3) = Val(Text15) Mod 256
'**********���Ƿ���
Dim BHtemp As String
Dim Btemp As Integer
Btemp = 0
Set rst = mconn.Execute("select yb_open from WTBdb where yb_id='" + Text15 + "'")
BHtemp = Trim(rst.Fields(0))
For i = 4 To 1 Step -1
If Left(Right(BHtemp, i), 1) Then
    Btemp = Btemp + 2 ^ (i - 1)
End If
Next i
rst.Close
'��ˮ��Ϣ   4,5- 6,7- 8,9 -10,11
Dim buyShu As Integer
If JTYes Then       '�Ƿ�ͨ����ˮ��
    For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '�����ô˱�
        Para(4 + (4 - i) * 2) = Val(Text4(4 - i)) Mod 256   '��ˮ����� ��λ��ǰ
        Para(5 + (4 - i) * 2) = Val(Text4(4 - i)) \ 256       '��ˮ����� ��λ�ں�
        Else
        Para(4 + (4 - i) * 2) = &H0                       '
        Para(5 + (4 - i) * 2) = &H0                       '
        End If
    Next i
Else
    For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '�����ô˱�
        buyShu = Val(Text4(4 - i)) * 100 '��ˮ������0.01��Ϊ��λ---����ͨ����ˮ�ۣ���ͨ����ǮΪ��λ
        Para(4 + (4 - i) * 2) = buyShu \ 100 Mod 256     '��ˮ����������
        Para(5 + (4 - i) * 2) = buyShu Mod 100           '��ˮ��С������
        Else
        Para(4 + (4 - i) * 2) = &H0                       '
        Para(5 + (4 - i) * 2) = &H0                       '
        End If
    Next i
End If
'**************����12-13-14-15*************************
For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '�����ô˱�
    Para(12 + (4 - i)) = Val(BUYcushu)           '����
    Else
    Para(12 + (4 - i)) = &H0                     '����
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
'����Ϊ1ʱ������
If Val(BUYcushu) = 1 Then
Para(22) = &H11
Else
Para(22) = &H0
End If
'����ʱ���־23
Para(23) = &H0
'������־24
Para(24) = &H11
'�հ�25-----��ͨ���ݱ�־
If JTYes Then
Para(25) = &H11
Else
Para(25) = &H0
End If
'��Ź�ˮ��־26 ????Ҫ��������124�����ݿ�Ϊ1101,����Ϊ1011?
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
'***************************************************************************************************************************************************************************************************************************************************************00
'*************��������************************
'2��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����2���ж�����*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("������λ�������")
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

End If          'disf

If DisD Then
'***************��������***********************************
'**************�����û�������******************************
Dim TempGZ As String, TempGD As String       '���ش���
Dim TempTZ As String        '͸֧��
Set rst = mconn.Execute("select * from wtddb where DS_name=(select top 1 yb_type from wtbddb where yb_id='" + Text15 + "')")
TempGZ = rst.Fields("ds_gznum")
TempTZ = rst.Fields("ds_tz")
rst.Close

'����0��
st = ser_102(icdev, 0, 2, 8)
If st < 0 Then
    MsgBox "����ʧ�ܣ���"
    Exit Sub
End If
'д�û����************************
Para(0) = &H0
Para(1) = &H0
Call ToBCD(Left(Right(Text15, 4), 2), Para(2))
Call ToBCD(Right(Text15, 2), Para(3))
'������־**************************
Para(4) = &H41
st = swr_102_hex(icdev, 0, 2, 5, Para(0))
If st < 0 Then
  MsgBox ("д��������")
  Exit Sub
End If
''''''''''''''''''''''''''''
If Val(BuyCushuD) = 1 Then      '��������
    '����1��***************************
    st = ser_102(icdev, 1, 0, 22)
    If st < 0 Then
        MsgBox "����ʧ�ܣ���"
        Exit Sub
    End If
    
    '������****************************
    Para(5) = &HC2
    Para(6) = &HA9
    Set rst = mconn.Execute("select Apass from Sysdate")
    Apass = rst.Fields(0)
    st = asc_hex(Apass, Para(7), 2)
    If st < 0 Then
        MsgBox ("��ȡ�������")
        Exit Sub
    End If
    rst.Close
    
    '������
    TempGD = FormatString(Val(Text8), 4)
    Call ToBCD(Left(TempGD, 2), Para(9))
    Call ToBCD(Right(TempGD, 2), Para(10))
    '���ش���
    TempGZ = FormatString(Val(TempGZ), 2)
    Call ToBCD(TempGZ, Para(11))
    '͸֧��
    TempTZ = FormatString(Val(TempTZ), 2)
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
      MsgBox ("д��������")
      Exit Sub
    End If
    '1��������λ����
    '*************������λ��0,�˶�����ǰ���ܶ�Ӧ����1���ж�����*****
    st = clrrd_102(icdev, 1)
    If st < 0 Then
      MsgBox ("������λ�������")
      Exit Sub
    End If
    '*************����1����������Ϊ2cc1067d9435************************
    Dim pass(6) As Byte
    pass(0) = &H2C
    pass(1) = &HC1
    pass(2) = &H6
    pass(3) = &H7D
    pass(4) = &H94
    pass(5) = &H35
    st = wesc_102(icdev, 1, 6, pass(0))
    If st < 0 Then
        MsgBox ("���Ŀ�1�������������")
        Exit Sub
    End If
    
'''''''''''''''''''''''''''''''''
Else     '�ճ�����
For i = 0 To 4
    Para(i) = &HFF
Next i
'������
TempGD = FormatString(Val(Text8), 4)
Call ToBCD(Left(TempGD, 2), Para(5))
Call ToBCD(Right(TempGD, 2), Para(6))
'���ش���
TempGZ = FormatString(Val(TempGZ), 2)
Call ToBCD(TempGZ, Para(7))
'͸֧��
TempTZ = FormatString(Val(TempTZ), 2)
Call ToBCD(TempTZ, Para(8))
For i = 9 To 12
    Para(i) = &HFF
Next i
Dim TempShu As String
TempShu = FormatString(Val(BuyCushuD), 4)
Call ToBCD(Left(TempShu, 2), Para(13))
Call ToBCD(Right(TempShu, 2), Para(14))
For i = 15 To 19
    Para(i) = &HFF
Next i
'�忨1��1-19�ֽ�
st = ser_102(icdev, 1, 0, 22)
If st < 0 Then
    MsgBox ("����ʧ�ܣ�")
    Exit Sub
End If
'д��
st = swr_102_hex(icdev, 1, 1, 20, Para(0))
If st < 0 Then
  MsgBox ("д��ʧ�ܣ���")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'''''''''''''''''''''''''''''''''''
End If   'if disd
End If      '

'����޸�����
'*************��������************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("���Ŀ��������")
    Exit Sub
End If

'1��������λ����
'*************������λ��0,�˶�����ǰ���ܶ�Ӧ����1���ж�����*****
st = clrrd_102(icdev, 1)
If st < 0 Then
  MsgBox ("������λ�������")
  Exit Sub
End If




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

Private Sub Command3_Click()
On Error GoTo errhandle
'��䲹����
Set rst = mconn.Execute("select bkfee from sysdate")
Text7 = rst.Fields(0)
rst.Close
Set rst = mconn.Execute("select * from YHdb where y_id='" + Text10 + "'")
If rst.EOF Then
    MsgBox "û��������֤�������Ϣ����ȷ���Ƿ������������Ӵ��û���Ϣ��"
    Frame2(0).Enabled = False
    Command1.Enabled = False
    Text10.SetFocus
    Exit Sub
Else
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "С��" & Trim(rst.Fields("y_dong")) & "��" & Trim(rst.Fields("y_dy")) & "��Ԫ" & Trim(rst.Fields("y_hao")) & "��"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    Frame2(0).Enabled = True
    Command1.Enabled = True
End If
rst.Close
'�����һ�ι�ˮ��Ϣ
Set rst = mconn.Execute("select * from WTBdb where yb_id='" + Text15 + "'and yb_buyid=(select max(yb_buyid) from WTBdb where yb_id='" + Text15 + "')")
If Not rst.EOF Then
Text4(0) = rst.Fields("yb_w1")
Text4(1) = rst.Fields("yb_w2")
Text4(2) = rst.Fields("yb_w3")
Text4(3) = rst.Fields("yb_w4")
'�ۼ���
Text5(0) = rst.Fields("yb_tw1")
Text5(1) = rst.Fields("yb_tw2")
Text5(2) = rst.Fields("yb_tw3")
Text5(3) = rst.Fields("yb_tw4")
'����
BUYcushu = rst.Fields("yb_num")
'���
rst.Close
For i = 0 To 3
Text6(i) = Val(Text4(i)) * Val(Text1(i))
Next i
DisF = True
Else
MsgBox "���û�û���κι�ˮ��Ϣ"
DisF = False
End If

'�����һ�ι�����Ϣ
Set rst = mconn.Execute("select * from WTBDdb where yb_id='" + Text15 + "'and yb_buyid=(select max(yb_buyid) from WTBDdb where yb_id='" + Text15 + "')")
If Not rst.EOF Then
Text8 = rst.Fields("yb_dn")
BuyCushuD = rst.Fields("yb_num")
DisD = True
Else
MsgBox "���û�û���κι�����Ϣ"
DisD = False
End If
rst.Close


Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub
Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

'********************
'�ж���Щ��û�����ò��������ܹ�ˮ
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='1'")
If rst.Fields(0) = "          " Then
Frame2(6).Caption = "ˮ��һ��" & "(δ����)"
Else
Frame2(6).Caption = "ˮ��һ��" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(0) = rst1.Fields("w_price")
    Text2(0) = rst.Fields(0)
    Text3(0) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='2'")
If rst.Fields(0) = "          " Then
Frame2(1).Caption = "ˮ�����" & "(δ����)"
Else
Frame2(1).Caption = "ˮ�����" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(1) = rst1.Fields("w_price")
    Text2(1) = rst.Fields(0)
    Text3(1) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='3'")
If rst.Fields(0) = "          " Then
Frame2(2).Caption = "ˮ������" & "(δ����)"
Else
Frame2(2).Caption = "ˮ������" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(2) = rst1.Fields("w_price")
    Text2(2) = rst.Fields(0)
    Text3(2) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='4'")
If rst.Fields(0) = "          " Then
Frame2(3).Caption = "ˮ���ģ�" & "(δ����)"
Else
Frame2(3).Caption = "ˮ���ģ�" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(3) = rst1.Fields("w_price")
    Text2(3) = rst.Fields(0)
    Text3(3) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
End Sub


