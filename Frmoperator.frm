VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Frmoperator 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Աά��"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " ����Ա���� "
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
         Caption         =   " ����IC������Ȩ"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   2385
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ���ݱ��ݺͻָ�Ȩ"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   2385
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " ��Ϣ�޸�Ȩ"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox ChkOppower 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Ӫҵ����Ȩ"
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
      Begin VB.CommandButton cmd�ر� 
         Caption         =   "�ر�(&C)"
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd�޸� 
         Caption         =   "�޸�(&E)"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdˢ�� 
         Caption         =   "ˢ��(&R)"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdɾ�� 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd��� 
         Caption         =   "���(&A)"
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
         Caption         =   "����(&S)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "����(&C)"
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
Dim F_isNewuser As Boolean          '����OR�޸�
Dim OpNewno As String                  '�²���Ա����
Dim F_Op(5) As String         '����Ա��Ϣ�ֶ�

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
 Picadd.Visible = False
  Picmain.Visible = True
  '/* �г����в���Ա����
  Call ListOperator(LstOpname)
End Sub
Private Sub cmdˢ��_Click()     '��ˢ�¡���ť
  'ֻ�ж��û�Ӧ�ó�����Ҫ
    Call ListOperator(LstOpname)
End Sub
Private Sub cmd�ر�_Click()     '���رա���ť
  Unload Me
End Sub
Private Sub cmdCancel_Click()
    Picadd.Visible = False
    Picmain.Visible = True
    LstOpname.Enabled = True
    Frame1.Enabled = False
End Sub
Private Sub cmd���_Click()     '����ӡ���ť
On Error GoTo errhandle
Dim F_NewUserno As String
If gUsername = "admini" Then
    Picadd.Visible = True
    Picmain.Visible = False
    LstOpname.Enabled = False
    Frame1.Enabled = True
    F_isNewuser = True          '��"���"�²���Ա��־Ϊ�棨True��
    F_NewUserno = getnewopno()  '�����ݿ������Ԥ����һ���µĲ���Ա��
    txtFields(0).Text = F_NewUserno   '�²���Ա����
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    ChkOppower(0).Value = 1     'Ĭ��:��ӵ�г�Ʊ����Ȩ
    ChkOppower(1).Value = 0     'Ĭ��:��ӵ����˰֤���á�����Ȩ
    ChkOppower(2).Value = 0     'Ĭ�ϣ���ӵ�����ݱ��ݺͻָ�Ȩ
    ChkOppower(3).Value = 0     'Ĭ�ϣ���ӵ�г�Ʊ���Ǽ�Ȩ
'    txtFields(1).SetFocus
Else
MsgBox "ֻ�г�������Ա���ܸ��Ĳ���Ա��Ϣ����"
End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
End Sub
Private Sub cmd�޸�_Click()     '���޸ġ���ť
On Error GoTo errhandle
    If gUsername = "admini" Then
    Dim m_Opname As String        '����Ա����
    Picadd.Visible = True
    Picmain.Visible = False
    LstOpname.Enabled = False
    Frame1.Enabled = True

        m_Opname = Trim$(LstOpname.Text)
        '/* 1���޸Ķ����Ƿ�Ϊ����������Ա����
        If m_Opname = "admini" Then
            '/* 1.1���޸Ķ����ǡ���������Ա��
            '/* 1.1.1���жϵ�ǰ��¼���û��Ƿ�Ϊ����������Ա����(ֻ�С���������Ա���ſ��޸��Լ�����Ϣ)
            If gUsername = "admini" Then
                '/* 1.1.1.1����¼�û��ǡ���������Ա��
                '/* ��Ȩ���޸ģ�������������Ա��������Ȩ�޼��������޸�
                txtFields(1).Enabled = False    '/* ��Txtfields��1���ı���Enabled����Ϊ�棨True��
                                                '/* ��Ӧ������ԱԱ�������ֶ�
'                CboPower.Enabled = False        '/* ��Ȩ�޼�����Ͽ�Enabled����Ϊ�棨True��
                                                '/* ��Ӧ��Ȩ�޼����ֶ�
                Call LstOpname_Click
            Else
                '/* 1.1.1.2���ǡ���������Ա�������޸ġ���������Ա������Ϣ
                Beep
                MsgBox "��û��Ȩ���޸ĳ�������Ա����Ϣ��", vbCritical + vbOKOnly, App.Title
                Picadd.Visible = False
                Picmain.Visible = True
                LstOpname.Enabled = True
                Exit Sub
            End If
        End If
    Else
    MsgBox "ֻ�г�������Ա���ܸ��Ĳ���Ա��Ϣ����"
    End If
    
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next

End Sub
Private Sub cmdsave_Click()     '�����桱��ť
On Error GoTo errhandle
    Dim sqlstr As String
    Dim m_Opname As String        '����Ա����
    m_Opname = Trim$(txtFields(1).Text)
    '/* 1������Ա�����Ƿ�Ϊ�գ�
    If m_Opname <> "" Then
        '/* 1.1�� ����Ա������Ϊ�ա������,�˳�
        
        '/* 1.1.1�� ��ϲ���Ա�Ĳ���ְ��
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
        
        '/* 1.1.2���Ƿ�Ϊ����²���Ա��
        If F_isNewuser Then
            '/* 1.1.2.1���ǡ���ӡ��²���Ա
            '/* �����²���Ա

            Dim scr As String       '����
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
'            '/* ��1�������ô洢���̣�insoperator���������ݿ�
'            mconn.Execute sqlstr, dbSQLPassThrough
            mconn.Execute ("insert into operator(operatorno,name,password,power) values ('" + Trim$(txtFields(0).Text) + "','" + Trim$(txtFields(1).Text) + "','" + Trim$(txtFields(2).Text) + "','" + m_Oppower + "')")
            '/* ��2������ListBox�м����²���Ա����
            LstOpname.AddItem m_Opname
            LstOpname.ListIndex = LstOpname.ListCount - 1
                  
            '/* ��4����������²���Ա��־Ϊ��(False)
            F_isNewuser = False

    
            If MsgBox("�µĲ���Ա�Ѳ��룡" + Chr$(13) + "��������²���Ա��", vbOKCancel + vbQuestion, App.Title) = vbOK Then
                Call cmd���_Click
                Exit Sub
            End If
            Frame1.Enabled = False
            Picadd.Visible = False
            Picmain.Visible = True
            LstOpname.Enabled = True
        Else
        '/* 1.1.2.2���ǡ���ӡ��²���Ա
        '/* ���²���Ա
            '/* ��1�����������ݿ�
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
            '/* ��2��������ListBox��Ӧ����Ա����
            If Trim$(LstOpname.Text) <> m_Opname Then
                Dim m_Oldindex As Integer
                m_Oldindex = LstOpname.ListIndex
                LstOpname.RemoveItem LstOpname.ListIndex
                LstOpname.AddItem m_Opname, m_Oldindex
                LstOpname.ListIndex = m_Oldindex
            End If
            MsgBox "����Ա��Ϣ�Ѹ��£�", vbOKOnly + vbInformation, App.Title
            Frame1.Enabled = False
            Picadd.Visible = False
            Picmain.Visible = True
            LstOpname.Enabled = True
        End If
  
    Else
        MsgBox "����Ա��������Ϊ�գ�", vbOKOnly + vbCritical, App.Title
        txtFields(1).SetFocus
        Exit Sub
    End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    'Resume Next

End Sub
Private Sub cmdɾ��_Click()     '��ɾ������ť
On Error GoTo errhandle
    Dim m_Operatorno As String * 3  '����Ա����
    Dim m_Opname As String          '����Ա����

    m_Opname = Trim$(LstOpname.Text)
    m_Operatorno = txtFields(0).Text

    '/* 1�����Ҫɾ���Ĳ���Ա�Ƿ�����ʹ��ϵͳ
'    If Not IsUseSystem(m_Operatorno) Then
    If gUsername <> m_Opname Then
    '/* 2.1��û��ʹ��ϵͳ
        '/* 2.1.1����ɾ���Ĳ���Ա�Ƿ�Ϊ����������Ա����
        If m_Opname = "admini" Then
            '/* 2.1.1.1���ǡ���������Ա��
            '/* ��ʾ������ɾ��
            Beep
            MsgBox "��������Ա(admini)����ɾ����", vbOKOnly + vbCritical, App.Title
        Else
            '/* 2.1.1.1�����ǡ���������Ա��
            '/* 2.1.1.1.1��ȷ���Ƿ�ɾ����
            If MsgBox("ɾ�������ݽ����ɻָ�������" + Chr$(13) + "ȷ���Ƿ�ɾ����", vbQuestion + vbOKCancel, "����...") = vbOK Then

'                sqlstr = "delete from operator where operatorno='" + txtFields(0).Text + "'"
                mconn.Execute ("delete from operator where operatorno='" + txtFields(0).Text + "'")
                Dim m_Oldindex As Integer   '��ɾ���Ĳ���Աδɾ��֮ǰ��ListBox�����
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
    '/* 2.2��Ҫɾ�����ǵ�ǰ���ڵ�¼���û���
    '/* ��ʾ������ɾ��
        Beep
        MsgBox "�ò���Ա����ʹ��ϵͳ������ɾ����", vbOKOnly + vbCritical, App.Title
    End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next

End Sub

'/* ��ָ��ListBox����ʾ����Ա����
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
'* ������   getnewopno()
'* ���ܣ�   Ԥ������һ���²���Ա����
'* ��ڲ����� ��
'* ���أ�   һ���²���Ա����
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

Private Sub LstOpname_Click()   '/* ���ListBox

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
            txtFields(2) = Space(6) '��6���ո�
        End If
'        If Not IsNull(rst.Fields("power").Value) Then
'            txtFields(3) = Trim$(rst.Fields("power").Value)
'        Else
'            txtFields(3) = ""
'        End If
        
        '�б���ʾ����Ա�Ĳ���ְ��
        Dim m_Oppower As String
        If Not IsNull(rst.Fields("power").Value) Then
            m_Oppower = Trim$(rst.Fields("power").Value)
        Else
            m_Oppower = ""
        End If
        If InStr(m_Oppower, "A") <> 0 Then    'ӵ�г�Ʊ����Ȩ
            ChkOppower(0).Value = 1
        Else
            ChkOppower(0).Value = 0
        End If
        If InStr(m_Oppower, "B") <> 0 Then    'ӵ�в������Ȩ
            ChkOppower(1).Value = 1
        Else
            ChkOppower(1).Value = 0
        End If
        If InStr(m_Oppower, "C") <> 0 Then     'ӵ�����ݱ��ݺͻָ�Ȩ
            ChkOppower(2).Value = 1
        Else
            ChkOppower(2).Value = 0
        End If
        If InStr(m_Oppower, "D") <> 0 Then     'IC���ع�Ȩ
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

