VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��¼"
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
   StartUpPosition =   2  '��Ļ����
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
      ToolTipText     =   "�ڴ������û���"
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
      Caption         =   "ȷ��"
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
      Caption         =   "ȡ��"
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
      ToolTipText     =   "�ڴ������û�����"
      Top             =   840
      Width           =   2085
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000004&
      Caption         =   "�û�����(&U):"
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000004&
      Caption         =   "����(&P):"
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
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
Unload Me
Call QuitSystem
End Sub

Private Sub cmdOK_Click()

    Dim m_inputUser, m_inputPwd As String
    If Trim(Comboopname.Text) = "" Then
    MsgBox ("�û�������Ϊ�գ���")
    Comboopname.SetFocus
    Exit Sub
    End If
    m_inputUser = Trim$(Comboopname.Text)    '������û�����
    m_inputPwd = Trim$(txtpassword.Text)     '����Ŀ���
    Dim m_power As Integer  '����Ȩ�޼���
    Dim m_password As String    '����Ա����(�����ݿ��ж���)
    On Error GoTo errhand
    Set rst = mconn.Execute("select operatorno,password,power from operator where name='" + m_inputUser + "'")
    If Not IsNull(rst.Fields("password").Value) Then
        m_password = Trim$(rst.Fields("password").Value)
    Else
        m_password = ""
    End If
'    m_power = rst.Fields("power").Value
    
    '�����ȷ������
    If txtpassword = m_password Then
        LoginSucceeded = True
        Me.Hide
'-----------------------------------------------------------------------------
        gUsername = m_inputUser                           '��ȫ�ֲ���Ա����
        gUserno = Trim$(rst.Fields("operatorno").Value)   '��ȫ�ֲ���Ա����
        gPassword = m_inputPwd    '��ȫ���û�����
'        gUserpower = m_power   '��ȫ���û�Ȩ��
'        If IsNull(rst.Fields("op_power").Value) = False Then
'            gUserOpFun = Trim(rst.Fields("op_power").Value) '��ȫ�ֲ���Աְ��
'        Else
'            gUserOpFun = ""
'        End If
    rst.Close

'----------------------------------------------------------------------------
      '���²�����¼
'        datein = Format(Now, "yyyy-MM-dd hh:mm:ss")
'        mconn.Execute ("insert into oprecord (operatorno,date_in) values('" + gUserno + "','" + datein + "') ")
    Else
        MsgBox "��Ч�����룬������!", , "��¼"
        txtpassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
'    Unload Me
    Exit Sub
errhand:
    MsgBox ("�û��������벻�ԣ����飡")
    Comboopname.SetFocus
End Sub

Private Sub Form_Load()
'���"����Ա����"��Ͽ�
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

Private Sub Form_Unload(Cancel As Integer)  'ж�ش���
Call QuitSystem
End Sub
