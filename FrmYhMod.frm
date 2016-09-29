VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmYhMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户检索"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "FrmYhMod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9510
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9120
      OleObjectBlob   =   "FrmYhMod.frx":030A
      Top             =   5880
   End
   Begin VB.CommandButton Command4 
      Caption         =   "购水信息"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取  消"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改信息"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   6
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "^用户编号|^姓名       |^身份证号         |^联系电话      |^详细住址                               |^备注         "
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "检索条件："
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton Command5 
         Caption         =   "刷新"
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "按详细住址查询"
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "确定"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3400
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "按电话号码查询"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按身份证号码查询"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "FrmYhMod.frx":053E
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmYhMod.frx":05B2
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmYhMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim m_QuerySQLstr As String

Private Sub Command1_Click()
YHMod = True
FrmyzAdd.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Option1.Value Then           '按身份证号码查询
    If Text1 = "" Then
    MsgBox "查询条件不能为空！"
    Text1.SetFocus
    Exit Sub
    End If
m_QuerySQLstr = "select * from YHdb where y_id='" + Text1 + "'"
Call RefreshYH(m_QuerySQLstr)
ElseIf Option2.Value Then           '按用户电话查询
    If Text1 = "" Then
    MsgBox "查询条件不能为空！"
    Text1.SetFocus
    Exit Sub
    End If
m_QuerySQLstr = "select * from YHdb where y_tel='" + Text1 + "'"
Call RefreshYH(m_QuerySQLstr)
ElseIf Option3.Value Then           '按用户地址查询
    If Text1 = "" Or Text2 = "" Or Text3 = "" Then
    MsgBox "查询条件不能为空！"
    Text1.SetFocus
    Exit Sub
    End If
If Text4 = "" Then
m_QuerySQLstr = "select * from YHdb where y_xq='" + Text1 + "'and y_dong='" + Text2 + "'and y_dy='" + Text3 + "'"
Else
m_QuerySQLstr = "select * from YHdb where y_xq='" + Text1 + "'and y_dong='" + Text2 + "'and y_dy='" + Text3 + "' and y_hao='" + Text4 + "'"
End If

Call RefreshYH(m_QuerySQLstr)
End If
End Sub

Private Sub Command4_Click()
YHModS = True
FrmBuyshuiP.Show vbModal
End Sub

Private Sub Command5_Click()
m_QuerySQLstr = "select * from YHdb"
Call RefreshYH(m_QuerySQLstr)
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
m_QuerySQLstr = "select * from YHdb"
Call RefreshYH(m_QuerySQLstr)

End Sub

Private Sub Form_Unload(Cancel As Integer)
YHMod = False
End Sub


Private Sub Option1_Click()
SkinLabel1.Caption = "请输入用户身份证号码："
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
SkinLabel2.Visible = False
End Sub
Private Sub Option2_Click()
SkinLabel1.Caption = "请输入用户电话号码："
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
SkinLabel2.Visible = False
End Sub
Private Sub Option3_Click()
SkinLabel1.Caption = "请输入用户详细地址："
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
SkinLabel2.Visible = True
End Sub

Sub RefreshYH(m_QuerySQLString As String)   '显示所有用户信息
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "现在没有任何信息！！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("y_no") + vbTab
                    dataitem = dataitem + .Fields("y_name") + vbTab
                    dataitem = dataitem + .Fields("y_id") + vbTab
                    dataitem = dataitem + .Fields("y_tel") + vbTab
                    dataitem = dataitem + Trim(.Fields("y_add")) + Trim(.Fields("y_xq")) + "小区" + Trim(.Fields("y_dong")) + "幢" + Trim(.Fields("y_dy")) + "单元" + Trim(.Fields("y_hao")) + "号" + vbTab
                    dataitem = dataitem + .Fields("y_memo") + vbTab
                 MSFlexGrid1.AddItem dataitem
                .MoveNext
            Wend
        End With
    End If
    MSFlexGrid1.Enabled = True
    rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

