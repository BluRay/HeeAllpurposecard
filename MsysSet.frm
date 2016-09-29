VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MsysSet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "水表详细参数设置"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "MsysSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10650
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "用水类型设置"
      TabPicture(0)   =   "MsysSet.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Command2(0)"
      Tab(0).Control(2)=   "Command1(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "水表类型设置"
      TabPicture(1)   =   "MsysSet.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1(1)"
      Tab(1).Control(1)=   "Command2(1)"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "用户水表布局"
      TabPicture(2)   =   "MsysSet.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command2(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command1(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "保存并退出"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   41
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取  消"
         Height          =   375
         Index           =   2
         Left            =   6600
         TabIndex        =   40
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Height          =   5295
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   9975
         Begin VB.Frame Frame4 
            Caption         =   "水表一"
            Height          =   2415
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   4455
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   0
               Left            =   480
               OleObjectBlob   =   "MsysSet.frx":035E
               TabIndex        =   46
               Top             =   960
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Index           =   0
               Left            =   240
               OleObjectBlob   =   "MsysSet.frx":03C4
               TabIndex        =   45
               Top             =   600
               Width           =   975
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   240
               Width           =   1815
            End
            Begin VB.ComboBox Combo2 
               Height          =   300
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Height          =   1215
               Index           =   0
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   960
               Width           =   2655
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Index           =   0
               Left            =   240
               OleObjectBlob   =   "MsysSet.frx":042C
               TabIndex        =   44
               Top             =   285
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "水表二"
            Height          =   2415
            Index           =   1
            Left            =   5040
            TabIndex        =   32
            Top             =   240
            Width           =   4695
            Begin VB.TextBox Text1 
               Height          =   1215
               Index           =   1
               Left            =   1320
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   960
               Width           =   2655
            End
            Begin VB.ComboBox Combo2 
               Height          =   300
               Index           =   1
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   600
               Width           =   1815
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               Index           =   1
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   240
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Index           =   1
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":0494
               TabIndex        =   47
               Top             =   360
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Index           =   1
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":04FC
               TabIndex        =   48
               Top             =   675
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   1
               Left            =   480
               OleObjectBlob   =   "MsysSet.frx":0564
               TabIndex        =   49
               Top             =   1035
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "水表三"
            Height          =   2415
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   2760
            Width           =   4455
            Begin VB.TextBox Text1 
               Height          =   1215
               Index           =   2
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   960
               Width           =   2535
            End
            Begin VB.ComboBox Combo2 
               Height          =   300
               Index           =   2
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   600
               Width           =   1815
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               Index           =   2
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   240
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Index           =   2
               Left            =   240
               OleObjectBlob   =   "MsysSet.frx":05D0
               TabIndex        =   50
               Top             =   360
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Index           =   2
               Left            =   240
               OleObjectBlob   =   "MsysSet.frx":0638
               TabIndex        =   51
               Top             =   675
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   2
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":06A0
               TabIndex        =   52
               Top             =   1035
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "水表四"
            Height          =   2415
            Index           =   3
            Left            =   5040
            TabIndex        =   24
            Top             =   2760
            Width           =   4695
            Begin VB.TextBox Text1 
               Height          =   1215
               Index           =   3
               Left            =   1320
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   27
               Top             =   960
               Width           =   2655
            End
            Begin VB.ComboBox Combo2 
               Height          =   300
               Index           =   3
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   600
               Width           =   1815
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               Index           =   3
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   240
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Index           =   3
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":070C
               TabIndex        =   53
               Top             =   360
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Index           =   3
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":0774
               TabIndex        =   54
               Top             =   675
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   3
               Left            =   480
               OleObjectBlob   =   "MsysSet.frx":07DC
               TabIndex        =   55
               Top             =   1035
               Width           =   735
            End
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "进入下一步"
         Height          =   375
         Index           =   1
         Left            =   -72960
         TabIndex        =   22
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取  消"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   -68400
         TabIndex        =   21
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "进入下一步"
         Height          =   375
         Index           =   0
         Left            =   -72960
         TabIndex        =   20
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取  消"
         Height          =   375
         Index           =   0
         Left            =   -68400
         TabIndex        =   19
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   9975
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   3135
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   5530
            _Version        =   393216
            FormatString    =   "^水表型号           |^水表采样点            "
         End
         Begin VB.Frame Frame6 
            Caption         =   "添加水表类型："
            Height          =   1575
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   9615
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "MsysSet.frx":0848
               TabIndex        =   57
               Top             =   480
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   480
               OleObjectBlob   =   "MsysSet.frx":08B2
               TabIndex        =   56
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton Command7 
               Caption         =   "保 存"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5160
               TabIndex        =   18
               Top             =   1080
               Width           =   1575
            End
            Begin VB.CommandButton Command6 
               Caption         =   "添 加"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2520
               TabIndex        =   17
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   6000
               TabIndex        =   15
               Top             =   360
               Width           =   2175
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1800
               TabIndex        =   14
               Top             =   360
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5295
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   9975
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2535
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4471
            _Version        =   393216
            Cols            =   6
            TextStyle       =   4
            FormatString    =   "^用水类型    |^    单价    |^   显示报警量 |^   关阀报警量 |^    限购水量 |^    设置时间   "
         End
         Begin VB.Frame Frame5 
            Caption         =   "添加用水类型设置："
            Height          =   2295
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   9495
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   6120
               TabIndex        =   9
               Top             =   1200
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   5160
               OleObjectBlob   =   "MsysSet.frx":091A
               TabIndex        =   63
               Top             =   1320
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   4680
               OleObjectBlob   =   "MsysSet.frx":0982
               TabIndex        =   59
               Top             =   360
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":09F0
               TabIndex        =   58
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton Command4 
               Caption         =   "保存类型"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5520
               TabIndex        =   10
               Top             =   1800
               Width           =   1575
            End
            Begin VB.CommandButton Command3 
               Caption         =   "重新填写"
               Height          =   375
               Left            =   1800
               TabIndex        =   11
               Top             =   1800
               Width           =   1575
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   6120
               TabIndex        =   5
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1440
               TabIndex        =   8
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1440
               TabIndex        =   6
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   6120
               TabIndex        =   7
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               Height          =   375
               Left            =   1440
               TabIndex        =   4
               Top             =   240
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   5160
               OleObjectBlob   =   "MsysSet.frx":0A58
               TabIndex        =   60
               Top             =   840
               Width           =   3615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":0B1A
               TabIndex        =   61
               Top             =   840
               Width           =   3375
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   360
               OleObjectBlob   =   "MsysSet.frx":0BDC
               TabIndex        =   62
               Top             =   1320
               Width           =   3615
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   43
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "MsysSet.frx":0C9E
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9840
      OleObjectBlob   =   "MsysSet.frx":1271
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   3960
      OleObjectBlob   =   "MsysSet.frx":14A5
      TabIndex        =   42
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "MsysSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim rst1 As Recordset
Dim i As Integer
Dim j As Integer
Dim m_QuerySQLstr As String

Private Sub Combo1_Click(Index As Integer)
Text1(Index) = ""
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo errhandle
Select Case Index
    Case 0
      SSTab1.Tab = 1
     Call SSTab1_Click(1)
    Case 1
      SSTab1.Tab = 2
     Call SSTab1_Click(2)
    Case 2      '保存各水表参数
    
    
    
    '先清空原有数据
    If Combo1(0) = "" Then
    MsgBox "水表一必须设置参数！！"
    Exit Sub
    End If
    
    
    If Combo1(1) = "" Then
    Text1(1) = "    "
    End If
    If Combo1(2) = "" Then
    Text1(2) = "    "
    End If
    If Combo1(3) = "" Then
    Text1(3) = "    "
    End If
    
    '水表类型必填
    If Combo2(0) = "" Then
    MsgBox "请选择水表类型！！"
    Exit Sub
    End If
    If Combo2(1) = "" Then
    MsgBox "请选择水表类型！！"
    Exit Sub
    End If
    If Combo2(2) = "" Then
    MsgBox "请选择水表类型！！"
    Exit Sub
    End If
    If Combo2(3) = "" Then
    MsgBox "请选择水表类型！！"
    Exit Sub
    End If
    mconn.Execute ("delete from WTSdb")
    For i = 0 To 3
    mconn.Execute ("insert into WTSdb (wt_no,wt_type,wt_stype,wt_add)values('" + Trim(Str(i + 1)) + "','" + Combo1(i) + "','" + Combo2(i) + "','" + Text1(i) + "')")
    Next i
    MsgBox "设置成功！"
    Unload Me
End Select
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command2_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command3_Click()
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = Str(Year(Now)) + Str(Month(Now)) + Str(Day(Now))
Text2.SetFocus
Command4.Enabled = True
End Sub

Private Sub Command4_Click()
On Error GoTo errhandle

If Command3.Enabled Then
If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
MsgBox "所有浅黄色项目为必填项，检查是否有漏填！"
Exit Sub
End If
mconn.Execute ("insert into WTYdb (w_name,w_max,w_warn1,w_warn2,w_price,w_stime)values('" + Text2 + "','" + Text3 + "','" + Text4 + "','" + Text5 + "','" + Text6 + "','" + Text7 + "')")
MsgBox "添加成功！"
MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "^用水类型    |^    单价    |^   显示报警量 |^   关阀报警量 |^    限购水量 |^    设置时间"
    m_QuerySQLstr = "select * from WTYdb"
    Call RefreshWtType(m_QuerySQLstr)
    
Else
    If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "所有浅黄色项目为必填项，检查是否有漏填！"
    Exit Sub
    End If
    mconn.Execute ("update WTYdb set w_max='" + Text3 + "',w_warn1='" + Text4 + "',w_warn2='" + Text5 + "',w_price='" + Text6 + "',w_stime='" + Text7 + "' where w_name='" + Text2 + "' ")
    MsgBox "修改成功！"
    MSFlexGrid1.Clear
        MSFlexGrid1.FormatString = "^用水类型    |^    单价    |^   显示报警量 |^   关阀报警量 |^    限购水量 |^    设置时间"
        m_QuerySQLstr = "select * from WTYdb"
        Call RefreshWtType(m_QuerySQLstr)
    
End If

Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command7_Click()                '保存水表类型
On Error GoTo errhandle
mconn.Execute ("insert into WTdb (wt_type,wt_chaiyan)values('" + Text8 + "','" + Text9 + "')")
MsgBox "添加成功！"
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
 SSTab1.Tab = 0
Call SSTab1_Click(0)

'填充水表combox
Set rst = mconn.Execute("select w_name from WTYdb")
Set rst1 = mconn.Execute("select wt_type from WTdb")

  Do While Not rst.EOF
    For i = 0 To 3
    Combo1(i).AddItem rst.Fields(0)
    Next i
    rst.MoveNext
  Loop
  Do While Not rst1.EOF
    For i = 0 To 3
    Combo2(i).AddItem rst1.Fields(0)
    Next i
    rst1.MoveNext
  Loop
rst.Close
rst1.Close
End Sub
Sub RefreshWtType(m_QuerySQLString As String)   '显示所有用水类型
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "现在没有任何信息！请添加！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("w_name") + vbTab
                    dataitem = dataitem + Str(.Fields("w_price")) + vbTab
                    dataitem = dataitem + Str(.Fields("w_warn1")) + vbTab
                    dataitem = dataitem + Str(.Fields("w_warn2")) + vbTab
                    dataitem = dataitem + Str(.Fields("w_max")) + vbTab
                    dataitem = dataitem + .Fields("w_stime") + vbTab
                 MSFlexGrid1.AddItem dataitem
                .MoveNext
            Wend
        End With
    End If
    rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub
Sub RefreshWType(m_QuerySQLString As String)   '显示所有水表类型
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid2.Clear
        MSFlexGrid2.Enabled = False
        Beep
        MsgBox "现在没有任何信息！请添加！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid2.Rows = 1
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("wt_type") + vbTab
                    dataitem = dataitem + Str(.Fields("wt_chaiyan")) + vbTab
                    MSFlexGrid2.AddItem dataitem

                .MoveNext
            Wend
        End With
    End If
    rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub


Private Sub MSFlexGrid1_Click()
'修改现有数据
'填充此项数据
MSFlexGrid1.Col = 0
Text2 = MSFlexGrid1.Text
MSFlexGrid1.Col = 1
Text6 = MSFlexGrid1.Text
MSFlexGrid1.Col = 2
Text4 = MSFlexGrid1.Text
MSFlexGrid1.Col = 3
Text5 = MSFlexGrid1.Text
MSFlexGrid1.Col = 4
Text3 = MSFlexGrid1.Text
MSFlexGrid1.Col = 5
Text7 = MSFlexGrid1.Text
Command3.Enabled = False
Command4.Enabled = True
Text6.Enabled = True
Text5.Enabled = True
Text4.Enabled = True
Text3.Enabled = True
    Text7.Text = Str(Year(Now)) & FormatString(Str(Month(Now)), 2) & FormatString(Str(Day(Now)), 2)

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 On Error GoTo errhandle
If SSTab1.Tab = 0 Then
'*********当前是否开通阶梯水价，如开通，此处单价不可改！以钱为单位****************

    m_QuerySQLstr = "select * from WTYdb"
    Call RefreshWtType(m_QuerySQLstr)
    
    If JTYes Then
    Text6 = "0"
    Text6.Enabled = False
    Frame5.Caption = "添加用水类型设置：   注：当前已经开通阶梯水价！"
    SkinLabel10.Caption = "显示报警量：                                           元"
    SkinLabel9.Caption = "限购水量：                                            元"
    SkinLabel11.Caption = "关阀报警量：                                           元"
    End If
    
    
    
ElseIf SSTab1.Tab = 1 Then
    m_QuerySQLstr = "select * from WTdb"
    Call RefreshWType(m_QuerySQLstr)
ElseIf SSTab1.Tab = 2 Then
For i = 0 To 3
Combo1(i).Clear
Combo2(i).Clear
Next i
    '填充水表combox
Set rst = mconn.Execute("select w_name from WTYdb")
Set rst1 = mconn.Execute("select wt_type from WTdb")
  Do While Not rst.EOF
    For i = 0 To 3
    Combo1(i).AddItem rst.Fields(0)
    Next i
    rst.MoveNext
  Loop
  Do While Not rst1.EOF
    For i = 0 To 3
    Combo2(i).AddItem rst1.Fields(0)
    Next i
    rst1.MoveNext
  Loop
rst.Close
rst1.Close

    




    Set rst = mconn.Execute("select * from WTSdb")
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
    Exit Sub
    Else
        Set rst1 = mconn.Execute("select * from WTSdb where wt_no='1'")
        Text1(0) = " 原参数：" & rst1.Fields("wt_type") & "        备注是：" & rst1.Fields("wt_add")
        rst1.Close
        Set rst1 = mconn.Execute("select * from WTSdb where wt_no='2'")
        If Not rst1.BOF Then
        Text1(1) = " 原参数：" & rst1.Fields("wt_type") & "        备注是：" & rst1.Fields("wt_add")
        rst1.Close
        End If
        Set rst1 = mconn.Execute("select * from WTSdb where wt_no='3'")
        If Not rst1.BOF Then
        Text1(2) = " 原参数：" & rst1.Fields("wt_type") & "        备注是：" & rst1.Fields("wt_add")
        rst1.Close
        End If
        Set rst1 = mconn.Execute("select * from WTSdb where wt_no='4'")
        If Not rst1.BOF Then
        Text1(3) = " 原参数：" & rst1.Fields("wt_type") & "        备注是：" & rst1.Fields("wt_add")
        rst1.Close
        End If
    End If

    rst.Close
End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

