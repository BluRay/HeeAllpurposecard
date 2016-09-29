VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmyzAdd 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "业主信息登记"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   Icon            =   "FrmyzAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9150
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8520
      OleObjectBlob   =   "FrmyzAdd.frx":030A
      Top             =   4080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取 消"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   8895
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmyzAdd.frx":053E
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmyzAdd.frx":05A6
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmyzAdd.frx":060E
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6960
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6080
         TabIndex        =   11
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5200
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5280
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5280
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmyzAdd.frx":067A
         TabIndex        =   15
         Top             =   1440
         Width           =   7455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "FrmyzAdd.frx":07DE
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmyzAdd.frx":084A
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   3248
      OleObjectBlob   =   "FrmyzAdd.frx":08B4
      TabIndex        =   13
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmyzAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim YHno As String

Private Sub Command1_Click()
On Error GoTo errhandle
If Text2 = "" Then
MsgBox "身份证号码不能为空！"
Text2.SetFocus
Exit Sub
End If

If Not YHMod Then
    '身份正必须唯一
    Set rst = mconn.Execute("select * from YHdb where y_id='" + Text2 + "'")
    If Not rst.EOF Then
    MsgBox "这个身份证号码已经存在,不能重复！"
    Text2.SetFocus
    Exit Sub
    End If
End If

If YHMod Then       '当为修改用户信息时，更新数据库
mconn.Execute ("update YHdb set y_name='" + Text3 + "',y_id='" + Text2 + "',y_tel='" + Text4 + "',y_add='" + Text5 + "',y_memo='" + Text6 + "',y_xq='" + Text7 + "',y_dong='" + Text8 + "',y_dy='" + Text9 + "',y_hao='" + Text10 + "' where y_no='" + YHno + "'")
MsgBox "修改成功"
Unload Me
Else
Set rst = mconn.Execute("insert into YHdb (y_no,y_name,y_id,y_tel,y_add,y_memo,y_xq,y_dong,y_dy,y_hao) values ('" + Text1 + "','" + Text3 + "','" + Text2 + "','" + Text4 + "','" + Text5 + "','" + Text6 + "','" + Text7 + "','" + Text8 + "','" + Text9 + "','" + Text10 + "')")
If MsgBox("添加成功！是否继续添加用户？", vbYesNo) = vbYes Then
    Set rst = mconn.Execute("select max(y_no) from YHdb")
        If rst.EOF Then
        Text1 = "0000001"
        Else
        Text1 = FormatString((Val(rst.Fields(0)) + 1), 7)
        End If
    rst.Close
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Else
Unload Me
End If
End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'要先设置参数才能继续
'必先设置好参数
Set rst = mconn.Execute("select count(wt_type) from wtsdb ")
If rst.Fields(0) = 0 Then
MsgBox "请先设置参数"
Command1.Enabled = False
Exit Sub
End If
rst.Close



'自动生成用户编号
If YHMod Then   '是否为修改用户信息
Me.Caption = "用户信息修改"
SkinLabel18 = "用户信息修改"
'填充要修改的用户原信息
    If FrmYhMod.MSFlexGrid1.Enabled Then
        '显示详细信息
        FrmYhMod.MSFlexGrid1.Col = 0
        YHno = Trim$(FrmYhMod.MSFlexGrid1.Text)
    Set rst = mconn.Execute("select * from YHdb where y_no='" + YHno + "'")
    If Not rst.EOF Then
    Text1 = YHno
    Text2 = rst.Fields("y_id")
    Text3 = rst.Fields("y_name")
    Text4 = rst.Fields("y_tel")
    Text5 = rst.Fields("y_add")
    Text6 = rst.Fields("y_memo")
    Text7 = rst.Fields("y_xq")
    Text8 = rst.Fields("y_dong")
    Text9 = rst.Fields("y_dy")
    Text10 = rst.Fields("y_hao")
    End If
    End If
ElseIf GYHcha Then  '由购水信息显示用户信息
Me.Caption = "用户详细信息"
SkinLabel18 = "用户详细信息"
'填充要修改的用户原信息
    If FrmBUYcha.MSFlexGrid1.Enabled Then
        '显示详细信息
        FrmBUYcha.MSFlexGrid1.Col = 1
        YHno = Trim$(FrmBUYcha.MSFlexGrid1.Text)
    Set rst = mconn.Execute("select * from YHdb where y_no='" + YHno + "'")
    If Not rst.EOF Then
    Text1 = YHno
    Text2 = rst.Fields("y_id")
    Text3 = rst.Fields("y_name")
    Text4 = rst.Fields("y_tel")
    Text5 = rst.Fields("y_add")
    Text6 = rst.Fields("y_memo")
    Text7 = rst.Fields("y_xq")
    Text8 = rst.Fields("y_dong")
    Text9 = rst.Fields("y_dy")
    Text10 = rst.Fields("y_hao")
    Command1.Enabled = False
    End If
    End If

Else
Set rst = mconn.Execute("select count(y_no) from YHdb")
    If rst.Fields(0) = 0 Then
    Text1 = "0000001"
    Else
    Text1 = FormatString((Val(rst.Fields(0)) + 1), 7)
    End If
rst.Close
End If
End Sub
