VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   2310
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6705
   FillStyle       =   0  'Solid
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1364.824
   ScaleMode       =   0  'User
   ScaleWidth      =   6295.63
   StartUpPosition =   2  '屏幕中心
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   2580
      TabIndex        =   8
      Top             =   0
      Width           =   2640
      Begin VB.Image Image1 
         Height          =   2325
         Left            =   0
         Picture         =   "frmLogin.frx":030A
         Top             =   0
         Width           =   2715
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   3000
      OleObjectBlob   =   "frmLogin.frx":1D21
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox Comboopname 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Text            =   "Comboopname"
      ToolTipText     =   "在此输入用户名"
      Top             =   360
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   2760
      OleObjectBlob   =   "frmLogin.frx":1D8D
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "frmLogin.frx":1DFD
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000011&
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   3120
      MaskColor       =   &H80000009&
      TabIndex        =   3
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "在此输入用户密码"
      Top             =   840
      Width           =   2085
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000004&
      Caption         =   "用户名称(&U):"
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000004&
      Caption         =   "密码(&P):"
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
Unload Me
Call QuitSystem
End Sub

Private Sub cmdOK_Click()

    Dim m_inputUser, m_inputPwd As String
    If Trim(Comboopname.Text) = "" Then
    MsgBox ("用户名不能为空！！")
    Comboopname.SetFocus
    Exit Sub
    End If
    m_inputUser = Trim$(Comboopname.Text)    '输入的用户姓名
    m_inputPwd = Trim$(txtpassword.Text)     '输入的口令
    Dim m_power As Integer  '操作权限级别
    Dim m_password As String    '操作员口令(从数据库中读出)
    On Error GoTo errhand
    Set rst = mconn.Execute("select operatorno,password,power from operator where name='" + m_inputUser + "'")
    If Not IsNull(rst.Fields("password").Value) Then
        m_password = Trim$(rst.Fields("password").Value)
    Else
        m_password = ""
    End If
'    m_power = rst.Fields("power").Value
    
    '检查正确的密码
    If txtpassword = m_password Then
        LoginSucceeded = True
        Me.Hide
'-----------------------------------------------------------------------------
        gUsername = m_inputUser                           '置全局操作员名称
        gUserno = Trim$(rst.Fields("operatorno").Value)   '置全局操作员代号
        gPassword = m_inputPwd    '置全局用户口令
'        gUserpower = m_power   '置全局用户权限
'        If IsNull(rst.Fields("op_power").Value) = False Then
'            gUserOpFun = Trim(rst.Fields("op_power").Value) '置全局操作员职能
'        Else
'            gUserOpFun = ""
'        End If
    rst.Close

'----------------------------------------------------------------------------
      '更新操作记录
'        datein = Format(Now, "yyyy-MM-dd hh:mm:ss")
'        mconn.Execute ("insert into oprecord (operatorno,date_in) values('" + gUserno + "','" + datein + "') ")
    Else
        MsgBox "无效的密码，请重试!", , "登录"
        txtpassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
'    Unload Me
    Exit Sub
errhand:
    MsgBox ("用户名或密码不对，请检查！")
    Comboopname.SetFocus
End Sub

Private Sub Form_Load()
'填充"操作员名称"组合框
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

Set rst = mconn.Execute("select name from operator")
Comboopname.Clear
            Do While Not rst.EOF
                Comboopname.AddItem rst.Fields(0).Value
                rst.MoveNext
            Loop
rst.Close
Comboopname = "admini"
End Sub

Private Sub Form_Unload(Cancel As Integer)  '卸载窗体
Call QuitSystem
End Sub
