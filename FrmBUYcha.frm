VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBUYcha 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ˮ��Ϣ��ѯ"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   Icon            =   "FrmBUYcha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11625
   StartUpPosition =   2  '��Ļ����
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FrmBUYcha.frx":030A
      Top             =   6720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ӡ"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�û���Ϣ"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   6600
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedCols       =   2
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ѯ������"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         ToolTipText     =   "��ʾ���û����ǰ���0��ʡ��"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   6240
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���û���ţ�"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���û����֤���룺"
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ȷ��"
         Height          =   495
         Left            =   9240
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   8040
         TabIndex        =   11
         Text            =   "Combo3"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   7200
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   6000
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   4080
         TabIndex        =   8
         Text            =   "Combo3"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   3240
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����ֹʱ�䣺               ��       ��        ��       ��              ��       ��       ��"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   8655
      End
   End
End
Attribute VB_Name = "FrmBUYcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim i As Integer, j As Integer
Dim m_QuerySQLstr As String

Private Sub Check1_Click()
If Check1.Value Then
For i = 0 To 1
Combo1(i).Enabled = True
Combo2(i).Enabled = True
Combo3(i).Enabled = True
Next i
Else
For i = 0 To 1
Combo1(i).Enabled = False
Combo2(i).Enabled = False
Combo3(i).Enabled = False
Next i
End If
End Sub

Private Sub Check2_Click()
If Check2.Value Then
Text1.Enabled = True
Text1.SetFocus
Check3.Value = 0
Else
Text1.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value Then
Text2.Enabled = True
Text2.SetFocus
Check2.Value = 0
Else
Text2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
GYHcha = True
FrmyzAdd.Show vbModal
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim date1 As String, date2 As String
If SorD = "gmcxD" Then  '�����ѯor��ˮ��ѯ
    If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
    MsgBox "��ѡ���ѯ������"
    ElseIf Check1.Value = 1 And Check2.Value = 0 And Check3.Value = 0 Then '�����ڲ�ѯ
    date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
    date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper where datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuyD(m_QuerySQLstr)
    ElseIf Check1.Value = 0 And Check3.Value = 1 Then '����Ų�ѯ
        m_QuerySQLstr = "select * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper where yb_id='" + Text2 + "'"
        Call RefreshBuyD(m_QuerySQLstr)
    
    ElseIf Check1.Value = 0 And Check2.Value = 1 Then '�����֤�Ų�ѯ
    If Text1 = "" Then
    MsgBox ("������Ҫ��ѯ�����֤�ţ�")
    Exit Sub
    End If
        m_QuerySQLstr = "select * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper where yb_id=(select y_no from YHdb where y_id='" + Text1 + "')"
        Call RefreshBuyD(m_QuerySQLstr)
    
    ElseIf Check1.Value = 1 And Check2.Value = 1 Then '����ͬʱ��ѯ
    If Text1 = "" Then
    MsgBox ("������Ҫ��ѯ�����֤�ţ�")
    Exit Sub
    End If
        
        date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
        date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper where yb_id=(select y_no from YHdb where y_id='" + Text1 + "') and datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuyD(m_QuerySQLstr)
    ElseIf Check1.Value = 1 And Check3.Value = 1 Then '��ʱ��ӱ��
        date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
        date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper where yb_id='" + Text2 + "' and datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuyD(m_QuerySQLstr)
    End If
    
Else
    If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
    MsgBox "��ѡ���ѯ������"
    ElseIf Check1.Value = 1 And Check2.Value = 0 And Check3.Value = 0 Then '�����ڲ�ѯ
    date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
    date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBdb where datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuy(m_QuerySQLstr)
    ElseIf Check1.Value = 0 And Check3.Value = 1 Then '����Ų�ѯ
        m_QuerySQLstr = "select * from WTBdb where yb_id='" + Text2 + "'"
        Call RefreshBuy(m_QuerySQLstr)
    
    ElseIf Check1.Value = 0 And Check2.Value = 1 Then '�����֤�Ų�ѯ
    If Text1 = "" Then
    MsgBox ("������Ҫ��ѯ�����֤�ţ�")
    Exit Sub
    End If
        m_QuerySQLstr = "select * from WTBdb where yb_id=(select y_no from YHdb where y_id='" + Text1 + "')"
        Call RefreshBuy(m_QuerySQLstr)
    
    ElseIf Check1.Value = 1 And Check2.Value = 1 Then '����ͬʱ��ѯ
    If Text1 = "" Then
    MsgBox ("������Ҫ��ѯ�����֤�ţ�")
    Exit Sub
    End If
        date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
        date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBdb where yb_id=(select y_no from YHdb where y_id='" + Text1 + "') and datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuy(m_QuerySQLstr)
    ElseIf Check1.Value = 1 And Check3.Value = 1 Then '��ʱ��ӱ��
        date1 = Format(Str(Combo1(0)) & "-" & Str(Combo2(0)) & "-" & Str(Combo3(0)), "YYYY-MM-dd")
        date2 = Format(Str(Combo1(1)) & "-" & Str(Combo2(1)) & "-" & Str(Combo3(1)), "YYYY-MM-dd")
        m_QuerySQLstr = "select * from WTBdb where yb_id='" + Text2 + "' and datediff('d',yb_date,'" + date1 + "')<=0 and datediff('d',yb_date,'" + date2 + "')>=0"
        Call RefreshBuy(m_QuerySQLstr)
    End If
End If
'*****************************************************
    If Not MSFlexGrid1.Enabled Then
        MSFlexGrid1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
'������¿�
For i = 0 To 1
    For j = 1 To 12
    Combo2(i).AddItem j
    Next j
    For j = 1 To 30
    Combo3(i).AddItem j
    Next j
Combo1(i) = Year(Now)
Combo2(i) = Month(Now)
Combo3(i) = Day(Now)
    For j = 1 To 4
    Combo1(i).AddItem (Val(Year(Now) - j))
    Next j
Next i


'Ĭ����䱾�����й�ˮ/����Ϣ
If SorD = "gmcxD" Then
Me.Caption = "�û������ѯ"
m_QuerySQLstr = "select top 50 * from WTBDdb left outer join operator on operator.operatorno=WTBDdb.yb_oper order by yb_date desc"
Call RefreshBuyD(m_QuerySQLstr)
Else
m_QuerySQLstr = "select top 50 * from WTBdb order by yb_date desc"
Call RefreshBuy(m_QuerySQLstr)
End If
End Sub

Sub RefreshBuy(m_QuerySQLString As String)   '��ʾ������ˮ����
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "û���κ���Ϣ��", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "^��ˮ���|^�û����|^��ˮ����            |^��ˮ����|^��ˮ���|^��һ������|^���������|^����������|^���Ĺ�����|^��һ�ܹ���|^����ܹ���|^�����ܹ���|^�����ܹ���"
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("yb_buyid") + vbTab
                    dataitem = dataitem + FormatString(Str(.Fields("yb_id")), 7) + vbTab
                    dataitem = dataitem + Trim(.Fields("yb_date")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_num")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_money")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_w1")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_w2")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_w3")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_w4")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_tw1")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_tw2")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_tw3")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_tw4")) + vbTab
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
Sub RefreshBuyD(m_QuerySQLString As String)   '��ʾ�����õ�����
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "û���κ���Ϣ��", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "^������|^�û����|^��������            |^�������|^������  |^������  |^�ܹ���  |^��������     |^       ����Ա"
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("yb_buyid") + vbTab
                    dataitem = dataitem + FormatString(Val(.Fields("yb_id")), 7) + vbTab
                    dataitem = dataitem + Trim(.Fields("yb_date")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_num")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_money")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_dn")) + vbTab
                    dataitem = dataitem + Str(.Fields("yb_tdn")) + vbTab
                    dataitem = dataitem + .Fields("yb_type") + vbTab
                    
                    dataitem = dataitem + .Fields("name") + vbTab
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

Private Sub Form_Unload(Cancel As Integer)
GYHcha = False
End Sub

Private Sub MSFlexGrid1_DblClick()
Call Command1_Click
End Sub

Private Sub Text2_LostFocus()
Text2 = FormatString(Text2, 7)
End Sub
