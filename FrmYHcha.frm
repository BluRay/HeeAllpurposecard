VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmYHcha 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户分析"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11025
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmYHcha.frx":0000
      Top             =   6360
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打印"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "购水信息"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   6240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   6
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      FormatString    =   "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "查询条件："
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   9240
         TabIndex        =   12
         Text            =   "5"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   9240
         TabIndex        =   11
         Text            =   "6"
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "用户月均购电量小于指定量的用户          度/月"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3480
         TabIndex        =   4
         Text            =   "6"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   3480
         TabIndex        =   3
         Text            =   "5"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连续超过指定时间未购水的用户               个月"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   5175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "用户月均购水量小于指定量的用户             吨/月"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连续超过指定时间未购电的用户            个月"
         Height          =   255
         Left            =   6120
         TabIndex        =   13
         Top             =   720
         Width           =   4335
      End
   End
End
Attribute VB_Name = "FrmYHcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset, rst1 As Recordset
Dim idTemp As String
Dim dataitem As String
Dim m_QuerySQLstr As String

Private Sub Command1_Click()

On Error GoTo errhandle
If Option1.Value Then           '按月均用量少于指定值查询
Set rst1 = mconn.Execute("select yb_id ,sum(yb_w1)+sum(yb_w2)+sum(yb_w3)+sum(yb_w4),datediff('dd',min(yb_date),max(yb_date)) from wtbdb group by yb_id")
If Not rst1.BOF Then rst1.MoveFirst
    MSFlexGrid1.Rows = 1
  
    While Not rst1.EOF
        If rst1.Fields(2) <> 0 Then      '----只购水一次的用户以外
        If Val(rst1.Fields(1)) / Val(rst1.Fields(2)) * 30 < Val(Text1) Then
        idTemp = rst1.Fields(0)  '得到此用户ID
        
        Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
         If Not rst.BOF Then rst.MoveFirst
         If rst.EOF Then
         MSFlexGrid1.Clear
         MSFlexGrid1.Enabled = False
         Beep
         MsgBox "现在没有任何信息！！", vbOKOnly + vbInformation, App.Title
         Else
            With rst
            MSFlexGrid1.FormatString = "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
'            dataitem = ""
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
       rst.Close
       End If
       End If
       rst1.MoveNext
     Wend
     
rst1.Close


ElseIf Option2.Value Then       '按超过时间不购水查询

Set rst1 = mconn.Execute("select max(yb_date),yb_id from wtbdb group by yb_id ")
If Not rst1.BOF Then rst1.MoveFirst
    MSFlexGrid1.Rows = 1
  
    While Not rst1.EOF
        If DateDiff("m", CDate(rst1.Fields(0)), CDate(Now)) > Val(Text2) Then
        idTemp = rst1.Fields(1)  '得到此用户ID
        
        Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
         If Not rst.BOF Then rst.MoveFirst
         If rst.EOF Then
         MSFlexGrid1.Clear
         MSFlexGrid1.Enabled = False
         Beep
         MsgBox "现在没有任何信息！！", vbOKOnly + vbInformation, App.Title
         Else
            With rst
            MSFlexGrid1.FormatString = "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
'            dataitem = ""
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
       rst.Close
       End If
       rst1.MoveNext
     Wend
rst1.Close

ElseIf Option3.Value Then       '按超过时间不购水查询
'****************************************************
'Set rst1 = mconn.Execute("select yb_id ,sum(cast(yb_dn as int)),datediff('dd',min(yb_date),max(yb_date)) from wtbddb group by yb_id")
Set rst1 = mconn.Execute("select yb_id ,sum(yb_w1)+sum(yb_w2)+sum(yb_w3)+sum(yb_w4),datediff('dd',min(yb_date),max(yb_date)) from wtbdb group by yb_id")

If Not rst1.BOF Then rst1.MoveFirst
    MSFlexGrid1.Rows = 1
  
    While Not rst1.EOF
        If rst1.Fields(2) <> 0 Then      '----只购水一次的用户以外
        If Val(rst1.Fields(1)) / Val(rst1.Fields(2)) * 30 < Val(Text4) Then
        idTemp = rst1.Fields(0)  '得到此用户ID
        
        Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
         If Not rst.BOF Then rst.MoveFirst
         If rst.EOF Then
         MSFlexGrid1.Clear
         MSFlexGrid1.Enabled = False
         Beep
         MsgBox "现在没有任何信息！！", vbOKOnly + vbInformation, App.Title
         Else
            With rst
            MSFlexGrid1.FormatString = "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
'            dataitem = ""
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
       rst.Close
       End If
       End If
       rst1.MoveNext
     Wend
     
rst1.Close

'****************************************************
ElseIf Option4.Value Then       '按超过时间不购电查询
Set rst1 = mconn.Execute("select max(yb_date),yb_id from wtbddb group by yb_id ")
If Not rst1.BOF Then rst1.MoveFirst
    MSFlexGrid1.Rows = 1
  
    While Not rst1.EOF
        If DateDiff("m", CDate(rst1.Fields(0)), CDate(Now)) > Val(Text3) Then
        idTemp = rst1.Fields(1)  '得到此用户ID
        
        Set rst = mconn.Execute("select * from YHdb where y_no='" + idTemp + "'")
         If Not rst.BOF Then rst.MoveFirst
         If rst.EOF Then
         MSFlexGrid1.Clear
         MSFlexGrid1.Enabled = False
         Beep
         MsgBox "现在没有任何信息！！", vbOKOnly + vbInformation, App.Title
         Else
            With rst
            MSFlexGrid1.FormatString = "^用户编号|^姓名       |^身份证号             |^联系电话      |^详细住址                            |^备注         "
'            dataitem = ""
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
       rst.Close
       End If
       rst1.MoveNext
     Wend
rst1.Close

End If
    
    
    If Not MSFlexGrid1.Enabled Then
        MSFlexGrid1.Enabled = True
    End If

Exit Sub
errhandle:
MsgBox (Error(ErR))
End Sub

Private Sub Command2_Click()
'根据用户编号得到该用户所有购水信息
FrmBuyshuiP.Show vbModal
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Option1_Click()
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub
Private Sub Option2_Click()
Text1.Enabled = False
Text2.Enabled = True
Text3.Enabled = False
Text4.Enabled = False
End Sub
Private Sub Option3_Click()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = True
End Sub
Private Sub Option4_Click()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = True
Text4.Enabled = False
End Sub
Private Sub MSFlexGrid1_DblClick()
Call Command2_Click
End Sub

