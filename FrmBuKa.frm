VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmBuKa 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户补卡"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text8 
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
         Name            =   "宋体"
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
      Caption         =   "取消"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "补卡"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "用户信息："
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
         Caption         =   "确 定"
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
      Caption         =   "用户最后一次购水信息："
      Enabled         =   0   'False
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9015
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "水表二："
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
         Caption         =   "水表三："
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
         Caption         =   "水表四："
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
         Caption         =   "水表一："
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
               Name            =   "宋体"
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
               Name            =   "宋体"
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
Dim Para(59) As Byte   '参数数组，共60字节
Dim BUYcushu As String, BuyCushuD As String
Dim DisF As Boolean, DisD As Boolean

'?????  补卡时购电数据？？？？？

Private Sub Command1_Click()
On Error GoTo errhandle
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
oldpass = "f0f0"
st = asc_hex(oldpass, password(0), 2)
st = csc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("核对IC卡密码错，请使用新卡或先回收卡！")
    Exit Sub
End If
'**************擦除0区******************************
st = ser_102(icdev, 0, 18, 5)
If st < 0 Then
    MsgBox ("擦卡出错！！")
    Exit Sub
End If
'写本系统卡标志
Para(0) = &H98
st = swr_102_hex(icdev, 0, 21, 1, Para(0))
If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If
''*************写用户卡标志***********************
Para(0) = &H10
st = swr_102_hex(icdev, 0, 18, 1, Para(0))
If st < 0 Then
MsgBox "写卡失败！"
Exit Sub
End If

If DisF Then    '是否有购电数据
'**************生成用户卡参数******************************
'区码0、1
Dim rst As Recordset
Set rst = mconn.Execute("select area from Sysdate")
st = asc_hex(rst.Fields(0), Para(0), 2)
rst.Close
'用户编号2、3
Para(2) = Val(Text15) / 256
Para(3) = Val(Text15) Mod 256
'**********表是否开启
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
'购水信息   4,5- 6,7- 8,9 -10,11
Dim buyShu As Integer
If JTYes Then       '是否开通阶梯水价
    For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '如启用此表
        Para(4 + (4 - i) * 2) = Val(Text4(4 - i)) Mod 256   '购水量金额 低位在前
        Para(5 + (4 - i) * 2) = Val(Text4(4 - i)) \ 256       '购水量金额 高位在后
        Else
        Para(4 + (4 - i) * 2) = &H0                       '
        Para(5 + (4 - i) * 2) = &H0                       '
        End If
    Next i
Else
    For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '如启用此表
        buyShu = Val(Text4(4 - i)) * 100 '购水量，以0.01吨为单位---不开通价梯水价，开通后以钱为单位
        Para(4 + (4 - i) * 2) = buyShu \ 100 Mod 256     '购水量整数部分
        Para(5 + (4 - i) * 2) = buyShu Mod 100           '购水量小数部分
        Else
        Para(4 + (4 - i) * 2) = &H0                       '
        Para(5 + (4 - i) * 2) = &H0                       '
        End If
    Next i
End If
'**************次数12-13-14-15*************************
For i = 4 To 1 Step -1
    If Val(Right(Left(BHtemp, i), 1)) Then           '如启用此表
    Para(12 + (4 - i)) = Val(BUYcushu)           '次数
    Else
    Para(12 + (4 - i)) = &H0                     '次数
    End If
Next i
'***16-20设置时间参数****************
For i = 16 To 20
Para(i) = &H0
Next i
'购水进位21,如果开通阶梯水价，此字节置00
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
'开户标志22
'次数为1时即开户
If Val(BUYcushu) = 1 Then
Para(22) = &H11
Else
Para(22) = &H0
End If
'设置时间标志23
Para(23) = &H0
'补卡标志24
Para(24) = &H11
'空白25-----开通阶梯标志
If JTYes Then
Para(25) = &H11
Else
Para(25) = &H0
End If
'表号购水标志26 ????要反过来？124表数据库为1101,卡内为1011?
Para(26) = Btemp
'校验27
Para(27) = &H0
For i = 0 To 26
    Para(27) = Para(27) Xor Para(i)
Next i

'故障位及报表返回数据
For i = 28 To 59
    Para(i) = &HFF
Next i





'**************擦除地址******************************
st = ser_102(icdev, 2, 0, 60)
If st < 0 Then
    MsgBox ("擦卡失败！")
    Exit Sub
End If

Screen.MousePointer = vbHourglass

st = swr_102_hex(icdev, 2, 2, 20, Para(0))
If st < 0 Then
  MsgBox ("写卡失败！！")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
st = swr_102_hex(icdev, 2, 22, 40, Para(20))
If st < 0 Then
  MsgBox ("写卡失败！！")
  Screen.MousePointer = vbDefault
  Exit Sub
End If
  Screen.MousePointer = vbDefault
'***************************************************************************************************************************************************************************************************************************************************************00
'*************更改密码************************
'2区读保护位清零
'*************读保护位清0,核对密码前不能对应用区2进行读操作*****
st = clrrd_102(icdev, 2)
If st < 0 Then
  MsgBox ("读保护位清零出错！")
  Exit Sub
End If
'*********把用户密码写入测试区***************************
Dim Apass As String
Set rst = mconn.Execute("select Apass from Sysdate")
Apass = rst.Fields(0)
'Call ToBCD(Val(Left(Apass, 2)), password(0))
'Call ToBCD(Val(Right(Apass, 2)), password(1))
st = asc_hex(Apass, password(0), 2)
If st < 0 Then
    MsgBox ("读取卡密码错")
    Exit Sub
End If
rst.Close
st = ser_102(icdev, 2, 84, 2)
If st < 0 Then
    MsgBox ("擦卡失败！")
    Exit Sub
End If
st = swr_102_hex(icdev, 2, 84, 2, password(0))
If st < 0 Then
    MsgBox ("写测试区密码出错！")
    Exit Sub
End If

End If          'disf

If DisD Then
'***************购电数据***********************************
'**************生成用户卡参数******************************
Dim TempGZ As String, TempGD As String       '过载次数
Dim TempTZ As String        '透支量
Set rst = mconn.Execute("select * from wtddb where DS_name=(select top 1 yb_type from wtbddb where yb_id='" + Text15 + "')")
TempGZ = rst.Fields("ds_gznum")
TempTZ = rst.Fields("ds_tz")
rst.Close

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
''''''''''''''''''''''''''''
If Val(BuyCushuD) = 1 Then      '开户购电
    '擦除1区***************************
    st = ser_102(icdev, 1, 0, 22)
    If st < 0 Then
        MsgBox "擦卡失败！！"
        Exit Sub
    End If
    
    '卡密码****************************
    Para(5) = &HC2
    Para(6) = &HA9
    Set rst = mconn.Execute("select Apass from Sysdate")
    Apass = rst.Fields(0)
    st = asc_hex(Apass, Para(7), 2)
    If st < 0 Then
        MsgBox ("读取卡密码错")
        Exit Sub
    End If
    rst.Close
    
    '购电量
    TempGD = FormatString(Val(Text8), 4)
    Call ToBCD(Left(TempGD, 2), Para(9))
    Call ToBCD(Right(TempGD, 2), Para(10))
    '过载次数
    TempGZ = FormatString(Val(TempGZ), 2)
    Call ToBCD(TempGZ, Para(11))
    '透支量
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
    
'''''''''''''''''''''''''''''''''
Else     '日常购电
For i = 0 To 4
    Para(i) = &HFF
Next i
'购电量
TempGD = FormatString(Val(Text8), 4)
Call ToBCD(Left(TempGD, 2), Para(5))
Call ToBCD(Right(TempGD, 2), Para(6))
'过载次数
TempGZ = FormatString(Val(TempGZ), 2)
Call ToBCD(TempGZ, Para(7))
'透支量
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
'清卡1区1-19字节
st = ser_102(icdev, 1, 0, 22)
If st < 0 Then
    MsgBox ("擦卡失败！")
    Exit Sub
End If
'写卡
st = swr_102_hex(icdev, 1, 1, 20, Para(0))
If st < 0 Then
  MsgBox ("写卡失败！！")
    Screen.MousePointer = vbDefault
  Exit Sub
End If
'''''''''''''''''''''''''''''''''''
End If   'if disd
End If      '

'最后修改密码
'*************更改密码************************
password(0) = &H1B
password(1) = &H6C
st = wsc_102(icdev, 2, password(0))
If st < 0 Then
    MsgBox ("更改卡密码出错！")
    Exit Sub
End If

'1区读保护位清零
'*************读保护位清0,核对密码前不能对应用区1进行读操作*****
st = clrrd_102(icdev, 1)
If st < 0 Then
  MsgBox ("读保护位清零出错！")
  Exit Sub
End If




MsgBox "补卡成功！"
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
'填充补卡费
Set rst = mconn.Execute("select bkfee from sysdate")
Text7 = rst.Fields(0)
rst.Close
Set rst = mconn.Execute("select * from YHdb where y_id='" + Text10 + "'")
If rst.EOF Then
    MsgBox "没有这个身份证号码的信息，请确认是否输错，或重新添加此用户信息。"
    Frame2(0).Enabled = False
    Command1.Enabled = False
    Text10.SetFocus
    Exit Sub
Else
    Text11 = rst.Fields("y_name")
    Text12 = rst.Fields("y_tel")
    Text13 = Trim(rst.Fields("y_add")) & Trim(rst.Fields("y_xq")) & "小区" & Trim(rst.Fields("y_dong")) & "幢" & Trim(rst.Fields("y_dy")) & "单元" & Trim(rst.Fields("y_hao")) & "号"
    Text14 = rst.Fields("y_memo")
    Text15 = rst.Fields("y_no")
    Frame2(0).Enabled = True
    Command1.Enabled = True
End If
rst.Close
'填充上一次购水信息
Set rst = mconn.Execute("select * from WTBdb where yb_id='" + Text15 + "'and yb_buyid=(select max(yb_buyid) from WTBdb where yb_id='" + Text15 + "')")
If Not rst.EOF Then
Text4(0) = rst.Fields("yb_w1")
Text4(1) = rst.Fields("yb_w2")
Text4(2) = rst.Fields("yb_w3")
Text4(3) = rst.Fields("yb_w4")
'累计量
Text5(0) = rst.Fields("yb_tw1")
Text5(1) = rst.Fields("yb_tw2")
Text5(2) = rst.Fields("yb_tw3")
Text5(3) = rst.Fields("yb_tw4")
'次数
BUYcushu = rst.Fields("yb_num")
'金额
rst.Close
For i = 0 To 3
Text6(i) = Val(Text4(i)) * Val(Text1(i))
Next i
DisF = True
Else
MsgBox "此用户没有任何购水信息"
DisF = False
End If

'填充上一次购电信息
Set rst = mconn.Execute("select * from WTBDdb where yb_id='" + Text15 + "'and yb_buyid=(select max(yb_buyid) from WTBDdb where yb_id='" + Text15 + "')")
If Not rst.EOF Then
Text8 = rst.Fields("yb_dn")
BuyCushuD = rst.Fields("yb_num")
DisD = True
Else
MsgBox "此用户没有任何购电信息"
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
'判断哪些表没有设置参数，不能购水
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='1'")
If rst.Fields(0) = "          " Then
Frame2(6).Caption = "水表一：" & "(未设置)"
Else
Frame2(6).Caption = "水表一：" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(0) = rst1.Fields("w_price")
    Text2(0) = rst.Fields(0)
    Text3(0) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='2'")
If rst.Fields(0) = "          " Then
Frame2(1).Caption = "水表二：" & "(未设置)"
Else
Frame2(1).Caption = "水表二：" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(1) = rst1.Fields("w_price")
    Text2(1) = rst.Fields(0)
    Text3(1) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='3'")
If rst.Fields(0) = "          " Then
Frame2(2).Caption = "水表三：" & "(未设置)"
Else
Frame2(2).Caption = "水表三：" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(2) = rst1.Fields("w_price")
    Text2(2) = rst.Fields(0)
    Text3(2) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
Set rst = mconn.Execute("select wt_type,wt_add from WTSdb where wt_no='4'")
If rst.Fields(0) = "          " Then
Frame2(3).Caption = "水表四：" & "(未设置)"
Else
Frame2(3).Caption = "水表四：" & Trim(rst.Fields(1))
    Set rst1 = mconn.Execute("select * from WTYdb where w_name='" + rst.Fields(0) + "'")
    Text1(3) = rst1.Fields("w_price")
    Text2(3) = rst.Fields(0)
    Text3(3) = rst1.Fields("w_max")
    rst1.Close
End If
rst.Close
End Sub


