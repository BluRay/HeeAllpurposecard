VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmSysSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ��������"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "FrmSysSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7485
   StartUpPosition =   2  '��Ļ����
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6840
      OleObjectBlob   =   "FrmSysSet.frx":030A
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ ��"
      Height          =   495
      Left            =   2895
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   6255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmSysSet.frx":053E
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":05BA
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":0622
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":068A
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":06F2
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":075A
         TabIndex        =   9
         Top             =   2760
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "FrmSysSet.frx":0810
         TabIndex        =   10
         Top             =   2280
         Width           =   2895
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmSysSet.frx":08C6
      TabIndex        =   8
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmSysSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset

Private Sub Command1_Click()
On Error GoTo errhandle
'��ɾ����������
If Text1 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
If Text2 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
If Text3 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
If Text4 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
If Text5 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
If Text6 = "" Then
MsgBox "������Ϣ������д��������"
Exit Sub
End If
'�ɵ����������ɱ���������
Dim Apass As String
'''If Val(Text1) > 1112 Then
'''Apass = FormatString((Val(Text1) - 1111), 4)
'''Else
'''Apass = FormatString(Text1, 4)
'''End If
'����ʱ���뱣�ֲ��䣬��ʹ�����ʱ��������룬�޸ĺ��������Ϊ1b6c     ????
Apass = "1b6c"
mconn.Execute ("delete from sysdate")
mconn.Execute ("insert into sysdate(area,name,address,tel,khfee,bkfee,Apass)values('" + Text1 + "','" + Text2 + "','" + Text3 + "','" + Text4 + "','" + Text5 + "','" + Text6 + "','" + Apass + "')")
MsgBox "�������óɹ�"
Unload Me
Exit Sub
errhandle:
MsgBox (Error(ErR))
Call QuitSystem
End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
If SysMod Then
'���ԭ��Ϣ
Text1.Enabled = False
Set rst = mconn.Execute("select * from sysdate")
Text1 = rst.Fields("area")
Text2 = rst.Fields("name")
Text3 = rst.Fields("address")
Text4 = rst.Fields("tel")
Text5 = rst.Fields("khfee")
Text6 = rst.Fields("bkfee")
rst.Close
End If


End Sub

'Private Sub Form_Unload(Cancel As Integer)
'Call Command1_Click
'End Sub
