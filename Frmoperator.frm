VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Frmoperator 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作员维护"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "Frmoperator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7095
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 操作员资料 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   4155
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoperator.frx":030A
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "Frmoperator.frx":0372
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         DataField       =   "operatorno"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   0
         Left            =   3720
         MaxLength       =   6
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 功能IC卡管理权"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   2385
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 数据备份和恢复权"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   2385
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 信息修改权"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " 营业管理权"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "password"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   960
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "operator"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "Frmoperator.frx":03D8
      TabIndex        =   18
      Top             =   840
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6360
      OleObjectBlob   =   "Frmoperator.frx":0442
      Top             =   4920
   End
   Begin VB.ListBox LstOpname 
      Height          =   3840
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.PictureBox Picmain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton cmd关闭 
         Caption         =   "关闭(&C)"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd修改 
         Caption         =   "修改(&E)"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd刷新 
         Caption         =   "刷新(&R)"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd删除 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd添加 
         Caption         =   "添加(&A)"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picadd 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3240
      ScaleHeight     =   615
      ScaleWidth      =   2775
      TabIndex        =   7
      Top             =   600
      Width           =   2775
      Begin VB.CommandButton cmdsave 
         Caption         =   "保存(&S)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "放弃(&C)"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frmoperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim F_isNewuser As Boolean          '增加OR修改
Dim OpNewno As String                  '新操作员代号
Dim F_Op(5) As String         '操作员信息字段

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
 Picadd.Visible = False
  Picmain.Visible = True
  '/* 列出所有操作员姓名
  Call ListOperator(LstOpname)
End Sub
Private Sub cmd刷新_Click()     '“刷新”按钮
  '只有多用户应用程序需要
    Call ListOperator(LstOpname)
End Sub
Private Sub cmd关闭_Click()     '“关闭”按钮
  Unload Me
End Sub
Private Sub cmdCancel_Click()
    Picadd.Visible = False
    Picmain.Visible = True
    LstOpname.Enabled = True
    Frame1.Enabled = False
End Sub
Private Sub cmd添加_Click()     '“添加”按钮
On Error GoTo errhandle
Dim F_NewUserno As String
If gUsername = "admini" Then
    Picadd.Visible = True
    Picmain.Visible = False
    LstOpname.Enabled = False
    Frame1.Enabled = True
    F_isNewuser = True          '置"添加"新操作员标志为真（True）
    F_NewUserno = getnewopno()  '向数据库服务器预申请一个新的操作员号
    txtFields(0).Text = F_NewUserno   '新操作员代号
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    ChkOppower(0).Value = 1     '默认:不拥有车票出售权
    ChkOppower(1).Value = 0     '默认:不拥有完税证领用、发放权
    ChkOppower(2).Value = 0     '默认：不拥有数据备份和恢复权
    ChkOppower(3).Value = 0     '默认：不拥有车票入库登记权
'    txtFields(1).SetFocus
Else
MsgBox "只有超级管理员才能更改操作员信息！！"
End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub
Private Sub cmd修改_Click()     '“修改”按钮
On Error GoTo errhandle
    If gUsername = "admini" Then
    Dim m_Opname As String        '操作员姓名
    Picadd.Visible = True
    Picmain.Visible = False
    LstOpname.Enabled = False
    Frame1.Enabled = True

        m_Opname = Trim$(LstOpname.Text)
        '/* 1、修改对象是否为“超级管理员”？
        If m_Opname = "admini" Then
            '/* 1.1、修改对象是“超级管理员”
            '/* 1.1.1、判断当前登录的用户是否为“超级管理员”？(只有“超级管理员”才可修改自己的信息)
            If gUsername = "admini" Then
                '/* 1.1.1.1、登录用户是“超级管理员”
                '/* 有权利修改，但“超级管理员”姓名、权限级别不允许修改
                txtFields(1).Enabled = False    '/* 置Txtfields（1）文本框Enabled属性为真（True）
                                                '/* 对应“操作员员姓名”字段
'                CboPower.Enabled = False        '/* 置权限级别组合框Enabled属性为真（True）
                                                '/* 对应“权限级别”字段
                Call LstOpname_Click
            Else
                '/* 1.1.1.2、非“超级管理员”不可修改“超级管理员”的信息
                Beep
                MsgBox "您没有权利修改超级管理员的信息！", vbCritical + vbOKOnly, App.Title
                Picadd.Visible = False
                Picmain.Visible = True
                LstOpname.Enabled = True
                Exit Sub
            End If
        End If
    Else
    MsgBox "只有超级管理员才能更改操作员信息！！"
    End If
    
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next

End Sub
Private Sub cmdsave_Click()     '“保存”按钮
On Error GoTo errhandle
    Dim sqlstr As String
    Dim m_Opname As String        '操作员姓名
    m_Opname = Trim$(txtFields(1).Text)
    '/* 1、操作员姓名是否为空？
    If m_Opname <> "" Then
        '/* 1.1、 操作员姓名不为空。保存后,退出
        
        '/* 1.1.1、 组合操作员的操作职能
        Dim m_Oppower As String
        m_Oppower = ""
        If ChkOppower(0).Value = 1 Then
            m_Oppower = "A"
        End If
        If ChkOppower(1).Value = 1 Then
            m_Oppower = m_Oppower + "B"
        End If
        If ChkOppower(2).Value = 1 Then
            m_Oppower = m_Oppower + "C"
        End If
        If ChkOppower(3).Value = 1 Then
            m_Oppower = m_Oppower + "D"
        End If
        
        '/* 1.1.2、是否为添加新操作员？
        If F_isNewuser Then
            '/* 1.1.2.1、是“添加”新操作员
            '/* 插入新操作员

            Dim scr As String       '密码
            Dim i As Integer
'            sqlstr = "insert into operator(operatorno,name,password,power) values ("
'            For i = 0 To 2
'                If i = 2 Then
'                    sqlstr = sqlstr + "'" + Trim$(txtFields(i).Text) + "'"
'                Else
'                    sqlstr = sqlstr + "'" + Trim$(txtFields(i).Text) + "',"
'                End If
'            Next i
'            sqlstr = sqlstr + ",'" + m_Oppower + "')"
'            '/* （1）、调用存储过程（insoperator）插入数据库
'            mconn.Execute sqlstr, dbSQLPassThrough
            mconn.Execute ("insert into operator(operatorno,name,password,power) values ('" + Trim$(txtFields(0).Text) + "','" + Trim$(txtFields(1).Text) + "','" + Trim$(txtFields(2).Text) + "','" + m_Oppower + "')")
            '/* （2）、在ListBox中加入新操作员姓名
            LstOpname.AddItem m_Opname
            LstOpname.ListIndex = LstOpname.ListCount - 1
                  
            '/* （4）、置添加新操作员标志为假(False)
            F_isNewuser = False

    
            If MsgBox("新的操作员已插入！" + Chr$(13) + "继续添加新操作员？", vbOKCancel + vbQuestion, App.Title) = vbOK Then
                Call cmd添加_Click
                Exit Sub
            End If
            Frame1.Enabled = False
            Picadd.Visible = False
            Picmain.Visible = True
            LstOpname.Enabled = True
        Else
        '/* 1.1.2.2、非“添加”新操作员
        '/* 更新操作员
            '/* （1）、更新数据库
'                F_Op(0) = Trim$(txtFields(0).Text)
'                F_Op(1) = Trim$(txtFields(1).Text)
'                F_Op(2) = Trim$(txtFields(2).Text)
'                F_Op(3) = Trim$(txtFields(3).Text)
'
'                sqlstr = "update operator set operator='" + F_Op(1) + "',password='" + F_Op(2) + "'" _
'                & ",power='" + F_Op(3) + "',op_power='" + m_Oppower + "' " _
'                & " where operatorno='" + F_Op(0) + "'"
'                mconn.Execute sqlstr, dbSQLPassThrough
            mconn.Execute ("update operator set name='" + Trim$(txtFields(1).Text) + "',password='" + Trim$(txtFields(2).Text) + "',power='" + m_Oppower + "'where operatorno='" + Trim$(txtFields(0).Text) + "'")
            '/* （2）、更新ListBox对应操作员姓名
            If Trim$(LstOpname.Text) <> m_Opname Then
                Dim m_Oldindex As Integer
                m_Oldindex = LstOpname.ListIndex
                LstOpname.RemoveItem LstOpname.ListIndex
                LstOpname.AddItem m_Opname, m_Oldindex
                LstOpname.ListIndex = m_Oldindex
            End If
            MsgBox "操作员信息已更新！", vbOKOnly + vbInformation, App.Title
            Frame1.Enabled = False
            Picadd.Visible = False
            Picmain.Visible = True
            LstOpname.Enabled = True
        End If
  
    Else
        MsgBox "操作员姓名不能为空！", vbOKOnly + vbCritical, App.Title
        txtFields(1).SetFocus
        Exit Sub
    End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    'Resume Next

End Sub
Private Sub cmd删除_Click()     '“删除”按钮
On Error GoTo errhandle
    Dim m_Operatorno As String * 3  '操作员代号
    Dim m_Opname As String          '操作员姓名

    m_Opname = Trim$(LstOpname.Text)
    m_Operatorno = txtFields(0).Text

    '/* 1、检查要删除的操作员是否正在使用系统
'    If Not IsUseSystem(m_Operatorno) Then
    If gUsername <> m_Opname Then
    '/* 2.1、没有使用系统
        '/* 2.1.1、欲删除的操作员是否为“超级管理员”？
        If m_Opname = "admini" Then
            '/* 2.1.1.1、是“超级管理员”
            '/* 提示：不能删除
            Beep
            MsgBox "超级管理员(admini)不能删除！", vbOKOnly + vbCritical, App.Title
        Else
            '/* 2.1.1.1、不是“超级管理员”
            '/* 2.1.1.1.1、确认是否删除？
            If MsgBox("删除后数据将不可恢复！！！" + Chr$(13) + "确认是否删除？", vbQuestion + vbOKCancel, "警告...") = vbOK Then

'                sqlstr = "delete from operator where operatorno='" + txtFields(0).Text + "'"
                mconn.Execute ("delete from operator where operatorno='" + txtFields(0).Text + "'")
                Dim m_Oldindex As Integer   '已删除的操作员未删除之前在ListBox中序号
                m_Oldindex = LstOpname.ListIndex
                LstOpname.RemoveItem m_Oldindex
                If m_Oldindex <= LstOpname.ListCount - 1 Then
                    LstOpname.ListIndex = m_Oldindex
                Else
                    LstOpname.ListIndex = LstOpname.ListCount - 1
                End If
                Call LstOpname_Click
            End If
        End If
    Else
    '/* 2.2、要删除的是当前正在登录的用户，
    '/* 提示，不能删除
        Beep
        MsgBox "该操作员正在使用系统，不能删除！", vbOKOnly + vbCritical, App.Title
    End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next

End Sub

'/* 在指定ListBox中显示操作员姓名
Private Sub ListOperator(LstOpname As ListBox)
On Error GoTo errhandle
    Set rst = mconn.Execute("select name from Operator order by operatorno")
    If Not rst.EOF Then rst.MoveFirst
    LstOpname.Clear
    While Not rst.EOF
        LstOpname.AddItem rst.Fields("name").Value
        rst.MoveNext
    Wend
    If LstOpname.ListCount <> 0 Then
        LstOpname.ListIndex = 0
    End If
    rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next
End Sub
'****************************************************************
'* 函数：   getnewopno()
'* 功能：   预先申请一个新操作员代号
'* 入口参数： 无
'* 返回：   一个新操作员代号
'****************************************************************
Function getnewopno() As String
On Error GoTo errhandle
Dim tempop1 As String
Set rst = mconn.Execute("select count(operatorno) from operator")
tempop1 = "WMS" & FormatString(Str(rst.Fields(0).Value + 1), 3)
    getnewopno = tempop1
    rst.Close
Exit Function
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next
End Function

Private Sub LstOpname_Click()   '/* 点击ListBox

    Dim rst As Recordset

    Set rst = mconn.Execute("select * from operator where name='" + LstOpname.Text + "'")
    If Not rst.EOF Then
        If Not IsNull(rst.Fields("operatorno").Value) Then
            txtFields(0) = Trim$(rst.Fields("operatorno").Value)
        Else
            txtFields(0) = ""
        End If
        If Not IsNull(rst.Fields("name").Value) Then
            txtFields(1) = Trim$(rst.Fields("name").Value)
        Else
            txtFields(1) = ""
        End If
        If Not IsNull(rst.Fields("password").Value) Then
            txtFields(2) = rst.Fields("password").Value
        Else
            txtFields(2) = Space(6) '置6个空格
        End If
'        If Not IsNull(rst.Fields("power").Value) Then
'            txtFields(3) = Trim$(rst.Fields("power").Value)
'        Else
'            txtFields(3) = ""
'        End If
        
        '判别并显示操作员的操作职能
        Dim m_Oppower As String
        If Not IsNull(rst.Fields("power").Value) Then
            m_Oppower = Trim$(rst.Fields("power").Value)
        Else
            m_Oppower = ""
        End If
        If InStr(m_Oppower, "A") <> 0 Then    '拥有车票出售权
            ChkOppower(0).Value = 1
        Else
            ChkOppower(0).Value = 0
        End If
        If InStr(m_Oppower, "B") <> 0 Then    '拥有财务管理权
            ChkOppower(1).Value = 1
        Else
            ChkOppower(1).Value = 0
        End If
        If InStr(m_Oppower, "C") <> 0 Then     '拥有数据备份和恢复权
            ChkOppower(2).Value = 1
        Else
            ChkOppower(2).Value = 0
        End If
        If InStr(m_Oppower, "D") <> 0 Then     'IC卡控管权
            ChkOppower(3).Value = 1
        Else
            ChkOppower(3).Value = 0
        End If
    End If

    rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next

End Sub

