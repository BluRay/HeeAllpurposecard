VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIFrmSys 
   BackColor       =   &H00E0E0E0&
   Caption         =   "����IC������ˮ�����ϵͳV3.0"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8880
   Icon            =   "MDIFrmSys.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrmSys.frx":6F0C2
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1931
      ButtonWidth     =   1455
      ButtonHeight    =   1773
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�����û�"
            Key             =   "t1"
            Object.ToolTipText     =   "�����û���Ϣ"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "���򿪻�"
            Key             =   "t2"
            Object.ToolTipText     =   "������ˮ"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�ճ�����"
            Key             =   "t3"
            Object.ToolTipText     =   "�ճ���ˮ"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "��IC��"
            Key             =   "t4"
            Object.ToolTipText     =   "��ȡIC������Ϣ"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�û�����"
            Key             =   "t5"
            Object.ToolTipText     =   "�û���ѯ������ָ���û�"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�����ѯ"
            Key             =   "t6"
            Object.ToolTipText     =   "��ѯ��ˮ��Ϣ"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ϵͳ����"
            Key             =   "t7"
            Object.ToolTipText     =   "����ϵͳ����"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "����"
            Key             =   "t8"
            Object.ToolTipText     =   "��ʾ������Ϣ"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "�˳�ϵͳ"
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "�˳�"
            Key             =   "t9"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "MDIFrmSys.frx":7C71A
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5370
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4833
            MinWidth        =   4833
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7832
            MinWidth        =   7832
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7832
            MinWidth        =   7832
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   3422
            MinWidth        =   3422
            TextSave        =   "2009-8-1"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   3246
            MinWidth        =   3246
            TextSave        =   "17:10"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   9240
      Left            =   0
      ScaleHeight     =   9180
      ScaleWidth      =   8820
      TabIndex        =   2
      Top             =   1095
      Width           =   8880
      Begin VB.Image Image1 
         Height          =   7635
         Left            =   0
         Picture         =   "MDIFrmSys.frx":7C94E
         Stretch         =   -1  'True
         Top             =   720
         Width           =   15240
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":89F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":8DFFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":8FB4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":916A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":931F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":94D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":98C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":9CE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIFrmSys.frx":A1106
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu m_main 
      Caption         =   "�ճ�����"
      Begin VB.Menu m_yzadd 
         Caption         =   "�û���Ϣ�Ǽ�"
         HelpContextID   =   101
      End
      Begin VB.Menu m_yzmod 
         Caption         =   "�û���Ϣά��"
      End
      Begin VB.Menu m_tab 
         Caption         =   "-"
      End
      Begin VB.Menu m_oper 
         Caption         =   "����Ա����"
      End
      Begin VB.Menu m_psd 
         Caption         =   "����Ա�����޸�"
      End
      Begin VB.Menu m_backup 
         Caption         =   "���ݱ��ݻָ�"
      End
   End
   Begin VB.Menu m_yy 
      Caption         =   "Ӫҵҵ��"
      Begin VB.Menu m_yykh 
         Caption         =   "���򿪻�"
      End
      Begin VB.Menu m_rcgm 
         Caption         =   "�ճ�����"
      End
      Begin VB.Menu m_tab2 
         Caption         =   "-"
      End
      Begin VB.Menu m_readIC 
         Caption         =   "��IC����Ϣ"
      End
      Begin VB.Menu m_card 
         Caption         =   "���տ�"
      End
      Begin VB.Menu m_tab22 
         Caption         =   "-"
      End
      Begin VB.Menu m_yhbk 
         Caption         =   "�û�����"
      End
      Begin VB.Menu m_yhtg 
         Caption         =   "�û��˹�"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu m_pcard 
      Caption         =   "����IC��"
      Begin VB.Menu m_szcard 
         Caption         =   "�������ÿ�(ˮ)"
      End
      Begin VB.Menu m_qlcard 
         Caption         =   "�������㿨(ˮ)"
      End
      Begin VB.Menu m_cxcard 
         Caption         =   "������ѯ��(ˮ)"
      End
      Begin VB.Menu m_chushihua 
         Caption         =   "������ʼ����(ˮ)"
      End
      Begin VB.Menu mtabb4 
         Caption         =   "-"
      End
      Begin VB.Menu m_chushiD 
         Caption         =   "������ʼ����(��)"
      End
      Begin VB.Menu tab5 
         Caption         =   "-"
      End
      Begin VB.Menu m_settime 
         Caption         =   "����ˮ��ʱ��"
      End
   End
   Begin VB.Menu m_cha 
      Caption         =   "��ѯ���� "
      Begin VB.Menu m_yhcl 
         Caption         =   "�û�����"
      End
      Begin VB.Menu m_gscx 
         Caption         =   "�����ѯ"
      End
   End
   Begin VB.Menu m_baobiao 
      Caption         =   "�����ӡ "
      Begin VB.Menu m_yhb 
         Caption         =   "�û�����"
      End
      Begin VB.Menu m_yssb 
         Caption         =   "����ˮ����"
      End
      Begin VB.Menu tabb 
         Caption         =   "-"
      End
      Begin VB.Menu m_sdbb 
         Caption         =   "���۵籨��"
      End
   End
   Begin VB.Menu m_sys 
      Caption         =   "ϵͳά�� "
      Begin VB.Menu m_area 
         Caption         =   "�����汾����"
      End
      Begin VB.Menu m_sysst 
         Caption         =   "ϵͳ��������"
      End
      Begin VB.Menu m_jieti 
         Caption         =   "����ˮ��"
      End
      Begin VB.Menu m_skin 
         Caption         =   "����Ƥ��"
         Begin VB.Menu m_blue 
            Caption         =   "��ɫ����"
         End
         Begin VB.Menu m_green 
            Caption         =   "��ɫ����"
         End
         Begin VB.Menu m_media 
            Caption         =   "Media"
         End
         Begin VB.Menu m_zhe 
            Caption         =   "Zhelezo"
         End
         Begin VB.Menu m_mac 
            Caption         =   "ƻ������"
         End
      End
   End
   Begin VB.Menu m_help 
      Caption         =   "�������� "
      Begin VB.Menu m_abont 
         Caption         =   "����"
      End
      Begin VB.Menu m_heoper 
         Caption         =   "��ǰ����Ա��Ϣ"
      End
      Begin VB.Menu m_shelp 
         Caption         =   "�鿴����"
      End
   End
   Begin VB.Menu m_exit 
      Caption         =   "�˳�ϵͳ "
   End
End
Attribute VB_Name = "MDIFrmSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub m_abont_Click()
frmAbout.Show vbModal
End Sub

Private Sub m_area_Click()
SysMod = True
FrmSysSet.Show vbModal
End Sub

Private Sub m_backup_Click()
MsgBox "Ϊ��֤���ݰ�ȫ����ϵͳ�����Զ����ݻ��ơ�"
End Sub



Private Sub m_chushiD_Click()
FrmZeroD.Show vbModal
End Sub

Private Sub m_jieti_Click()
FrmJieTi.Show vbModal
End Sub

Private Sub m_sdbb_Click()
FrmBUYrptD.Show vbModal
End Sub

Private Sub m_settime_Click()
FrmTimeSet.Show vbModal
End Sub

'****************����Ƥ��****************
Private Sub m_blue_Click()
On Error Resume Next
FileCopy App.Path & "\skins\B-Studio.skn", App.Path & "\B-Studio.skn"
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
End Sub

Private Sub m_zhe_Click()       '����Ƥ��
On Error Resume Next
FileCopy App.Path & "\skins\Zhelezo.skn", App.Path & "\B-Studio.skn"
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd

End Sub
Private Sub m_green_Click()     '����Ƥ��
On Error Resume Next
FileCopy App.Path & "\skins\Green.skn", App.Path & "\B-Studio.skn"
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
End Sub
Private Sub m_media_Click()     '����Ƥ��
On Error Resume Next
FileCopy App.Path & "\skins\media.skn", App.Path & "\B-Studio.skn"
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
End Sub

Private Sub m_mac_Click()       '����Ƥ��
On Error Resume Next
Kill App.Path & "\B-Studio.skn"
FileCopy App.Path & "\skins\mac.skn", App.Path & "\B-Studio.skn"
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
End Sub

Private Sub m_heoper_Click()
AboutMe.Show vbModal
End Sub

Private Sub m_card_Click()
FrmQinCard.Show vbModal
End Sub

Private Sub m_chushihua_Click()
FrmChuShiHua.Show vbModal
End Sub

Private Sub m_cxcard_Click()
FrmChaCard.Show vbModal
End Sub

Private Sub m_exit_Click()
If MsgBox("ȷ��Ҫ�˳�ϵͳ��", vbYesNo) = vbYes Then
Call QuitSystem
End If
End Sub

Private Sub m_gscx_Click()
SorD = "gmcx"
FrmSorD.Show vbModal
'FrmBUYcha.Show vbModal
End Sub

Private Sub m_oper_Click()
Frmoperator.Show vbModal
End Sub

Private Sub m_psd_Click()
FrmchangPwd.Show vbModal
End Sub

Private Sub m_qlcard_Click()
FrmZero.Show vbModal
End Sub

Private Sub m_rcgm_Click()
SorD = "rcgm"
FrmSorD.Show vbModal
'FrmWTB.Show vbModal
End Sub

Private Sub m_readIC_Click()
FrmRdCd.Show vbModal
End Sub

Private Sub m_shelp_Click()
    CommonDialog1.CancelError = True
    'On Error GoTo ErrHandler
    '���� HelpCommand ����
    CommonDialog1.HelpCommand = cdlHelpForceFile
    'ָ�������ļ���
    CommonDialog1.HelpFile = App.Path & "\ICWT.hlp"
    '��ʾ Windows �������档
    CommonDialog1.ShowHelp
End Sub

Private Sub m_sysst_Click()
SorD = "sysset"
FrmSorD.Show vbModal
'FrmWTB.Show vbModal

'MsysSet.Show vbModal
End Sub

Private Sub m_szcard_Click()
FrmsetCd.Show vbModal
End Sub

Private Sub m_yhb_Click()
'MDIFrmSys.CrystalReport1.ReportFileName = App.Path + "\YH.rpt"
''CrystalReport1.RetrieveDataFiles
'MDIFrmSys.CrystalReport1.WindowState = crptMaximized
'MDIFrmSys.CrystalReport1.PrintReport
'DataReport1.Orientation = rptOrientLandscape    '�����ӡ
'DataReport1.Title = gCorpName + "����Ա��" + gUsername

DrpYHdb.Refresh
DrpYHdb.Show vbModal
End Sub

Private Sub m_yhbk_Click()
FrmBuKa.Show vbModal
End Sub

Private Sub m_yhcl_Click()
FrmYHcha.Show vbModal
End Sub

Private Sub m_yssb_Click()
FrmBUYrpt.Show vbModal
End Sub

Private Sub m_yykh_Click()
SorD = "yykh"
FrmSorD.Show vbModal

'FrmYhKh.Show vbModal
End Sub

Private Sub m_yzadd_Click()
FrmyzAdd.Show vbModal
End Sub

Private Sub m_yzmod_Click()
FrmYhMod.Show vbModal
End Sub

Private Sub MDIForm_Load()
On Error GoTo errhandle
Dim rst As Recordset
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
 ' Ӧ�ó�������.
    If App.PrevInstance = True Then
        MsgBox "�����Ѿ����У������ٴ�װ�أ�", vbOKOnly + vbInformation, App.Title
        Unload Me
        Exit Sub
    End If
 gbDBOpenFlag = OpenDatabaseX()
Dim dateStart As String
Dim dateLast As String
Dim numDays As Integer     '��ʹ������

    Dim tempstr1 As String * 256
    Dim tempstr2 As String * 256
    Dim templon1 As Long
    Dim templon2 As Long
    Dim GetVal As Long
'**********�ж�ϵͳ�Ƿ��״�����*******************
'**********�״����г�ʼ��ϵͳ����ʼ�����ݿ����--����ˮ�۱�--��ˮ��????
Set rst = mconn.Execute("select * from Sysdate")
If rst.EOF Then
'**********�״�����ʱ��ȡӲ�����ģ��������ݿ�***************
''''    Call GetVolumeInformation("C:\", tempstr1, 256, GetVal, templon1, templon2, tempstr2, 256)
''''    Miwen = Right(CStr(GetVal), 8)  '����Ӳ�����к�ȡ����
''''    Skey = power(Miwen, Smy)
''''    'mconn.Execute ("delete from regsys ")
''''    mconn.Execute ("insert into regsys (HDnum,Skey) values('" + Miwen + "','" + Skey + "')")
'**********���ð棬���״�ʹ��ʱ��������ݿ�***************
dateStart = Format(CDate(Now), "yyyy-MM-dd")
mconn.Execute ("insert into regsys (HDnum,Skey) values('" + dateStart + "','" + dateStart + "')")


'ϵͳע�ᴰ��show

'**********������ˮ�۱����һ������-������*******************
mconn.Execute ("insert into sysjt(jietiyesno) values ('no')")
'**********������ˮ�۱����һ������-������*******************

MsgBox "               ��ӭʹ�ã�����" + Chr(13) + "ϵͳ��⵽�����״����У���������ϵͳ����������"
FrmSysSet.Show vbModal
End If
rst.Close
'**********ϵͳע��****************************************

        Call GetVolumeInformation("C:\", tempstr1, 256, GetVal, templon1, templon2, tempstr2, 256)
    Miwen = Right(CStr(GetVal), 8)  '����Ӳ�����к�ȡ����
    Set rst = mconn.Execute("select 1 from regsys where hdnum='" + Miwen + "'")
    If rst.BOF Then
    '�緢��Ӳ��������ע��ʱ��ͬ��������ע��
    FrmRegsys.Show vbModal
    End If
    rst.Close

'**********���ð�ʹ�ô�������******************************
Dim shiyongcishu As Integer 'ʹ�ô���
Dim War As String
'''Set rst = mconn.Execute("select shiyongcishu from regsys where hdnum='" + Miwen + "'")
'''shiyongcishu = Val(rst.Fields("shiyongcishu"))
'''shiyongcishu = shiyongcishu + 1
'''mconn.Execute ("update regsys set shiyongcishu='" + Trim(shiyongcishu) + "'where hdnum='" + Miwen + "'")
'''If shiyongcishu < 100 Then
'''War = "��ӭ���ñ�ϵͳ����������" & Trim(shiyongcishu) & "��ʹ�ã�����������" & Trim(100 - shiyongcishu) & "��"
'''MsgBox (War)
'''Else
'''MsgBox ("��ϵͳʹ�������ѵ�������������ϵ��ȡ��ʽ�汾��")
''' Call QuitSystem
'''End If
'''rst.Close
'**********��ʱ�����ð棬һ��������******************************
'
'������ʹ��ʱ������һ��ʹ��ʱ��ȽϺ�������ݿ�
Set rst = mconn.Execute("select * from regsys")
dateStart = rst.Fields("HDnum")
dateLast = rst.Fields("Skey")
rst.Close
'�����ǰʱ��С��ϵͳ���һ��ʹ��ʱ�䣬˵���û���ʱ����ǰ���ˣ�ϵͳ�˳�
If DateDiff("d", CDate(Now), CDate(dateLast)) > 0 Then
MsgBox "ϵͳ��⵽��ǰʱ������ʵʱ�䲻��������ʱ�����ã�"
     Call QuitSystem    '�˳�ϵͳ
ElseIf DateDiff("d", CDate(dateStart), CDate(Now)) > 30 Then      '���һ��ʹ��ʱ�䲻�ܴ���30��
MsgBox "ϵͳ�������ѵ�������ϵͳ�ṩ����ϵ�Ի�ȡ��ʽ�棡"
     Call QuitSystem    '�˳�ϵͳ
Else    '���ɼ�������
numDays = DateDiff("d", CDate(dateStart), CDate(Now))
War = "��ӭ���ñ�ϵͳ����������" & Trim(numDays) & "��ʹ�ã�����������" & Trim(30 - numDays) & "��"
MsgBox (War)
mconn.Execute ("update regsys set Skey='" + Trim(Format(CDate(Now), "yyyy-MM-dd")) + "'where HDnum='" + dateStart + "'")
End If


'**********��ʼ���˿�****************************************
    commport = GetCommPort()
    ExitIC
    icdev = -1
If gbDBOpenFlag = False Then
     MsgBox ("�����ݿ�ʧ�ܣ��������ݿ����ã�")
     Call QuitSystem    '�˳�ϵͳ
End If


YHMod = False       '�û���Ϣ�޸ı�־
GYHcha = False
SysMod = False
YHModS = False
frmLogin.Show vbModal   '�������Ա��¼����
Show                     '/* ��ʾ������
'Picture1.Picture = LoadPicture(App.Path & "\bgpic.jpg")    '���뱳��ͼ
'����ԱȨ���趨
        '�б���ʾ����Ա�Ĳ���ְ��
        Set rst = mconn.Execute("select power from operator where operatorno='" + gUserno + "'")
        Dim m_Oppower As String
        If Not IsNull(rst.Fields("power").Value) Then
            m_Oppower = Trim$(rst.Fields("power").Value)
        Else
            m_Oppower = ""
        End If
        If InStr(m_Oppower, "A") <> 0 Then
            m_yykh.Enabled = True
            m_rcgm.Enabled = True
            m_yhbk.Enabled = True
            m_yhtg.Enabled = True
        Else
            m_yykh.Enabled = False
            m_rcgm.Enabled = False
            m_yhbk.Enabled = False
            m_yhtg.Enabled = False
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(4).Enabled = False
        End If
        If InStr(m_Oppower, "B") <> 0 Then
            m_yzmod.Enabled = True
            m_area.Enabled = True
            m_sysst.Enabled = True
            m_oper.Enabled = True
        Else
            m_yzmod.Enabled = False
            m_area.Enabled = False
            m_sysst.Enabled = False
            m_oper.Enabled = False
            Toolbar1.Buttons(10).Enabled = False
        End If
        If InStr(m_Oppower, "C") <> 0 Then
            m_sys.Enabled = True
        Else
            m_sys.Enabled = False
        End If
        If InStr(m_Oppower, "D") <> 0 Then
            m_pcard.Enabled = True
        Else
            m_pcard.Enabled = False
        End If
        rst.Close
        
'm_cxcard.Enabled = False
'm_chaD.Enabled = False
'm_settime.Enabled = False

        
StatusBar1.Panels(1) = "��ǰ����Ա��" & gUsername
Set rst = mconn.Execute("select name from sysdate")
StatusBar1.Panels(2) = rst.Fields(0)
rst.Close
'**********��ǰϵͳ�Ƿ��Ѿ���ͨ����ˮ��*********
Set rst = mconn.Execute("select jieTiyesno from sysjt")
If Trim(rst.Fields(0)) = "yes" Then
JTYes = True
Else
JTYes = False
End If
rst.Close
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title
    Resume Next
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call QuitSystem
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.key
   Case "t1"
        Call m_yzadd_Click
   Case "t2"
        Call m_yykh_Click
   Case "t3"
        Call m_rcgm_Click
   Case "t4"
        Call m_readIC_Click
   Case "t5"
        Call m_yhcl_Click
   Case "t6"
        Call m_gscx_Click
   Case "t7"
        Call m_sysst_Click
   Case "t8"
        Call m_shelp_Click
   Case "t9"
        Call m_exit_Click
End Select

End Sub
