VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmBUYrptD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "售电营业报表"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8265
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "打印"
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "打印条件："
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8055
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "打印指定用户购电信息"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "打印指定日期内购电信息"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "打印全部购电信息"
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   2880
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   3960
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmBUYrptD.frx":0000
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmBUYrptD.frx":0074
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   2280
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmBUYrptD.frx":0150
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   495
      Left            =   3120
      OleObjectBlob   =   "FrmBUYrptD.frx":09FD
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6960
      OleObjectBlob   =   "FrmBUYrptD.frx":0A5E
      Top             =   120
   End
End
Attribute VB_Name = "FrmBUYrptD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
Text2 = Year(Now)
Text3 = Year(Now)
Combo1 = Month(Now)
Combo2 = Month(Now)
For i = 1 To 12
Combo1.AddItem i
Combo2.AddItem i
Next i
End Sub
Private Sub Option1_Click()
Text1.Visible = True
SkinLabel2.Caption = "请输入用户身份证号码："
Text2.Visible = False
Text3.Visible = False
Combo1.Visible = False
Combo2.Visible = False
SkinLabel1.Visible = False
End Sub

Private Sub Option2_Click()
Text1.Visible = False
SkinLabel2.Caption = "请输入要打印信息的起止年月："
Text2.Visible = True
Text3.Visible = True
Combo1.Visible = True
Combo2.Visible = True
SkinLabel1.Visible = True
End Sub

Private Sub Command1_Click()
Dim SumMoney As Single

Unload DrpBuyDian
Unload DataEnvironment1





On Error GoTo errhandle
'根据查询条件将要印的数据存入打印表中
'首先清空打印表
mconn.Execute ("delete from WTBDPdb")
SumMoney = 0
If Option1.Value Then           '打印指定用户信息
    '由身份证号得到用户ID
    Set rst = mconn.Execute("select y_no from YHdb where y_id='" + Text1 + "'")
    If rst.EOF Then
        MsgBox "没有此用户，请检查输入是否的误！"
        Exit Sub
    End If
    Set rst1 = mconn.Execute("select * from WTBDdb where yb_id='" + rst.Fields(0) + "'")
    If Not rst1.BOF Then rst1.MoveFirst
        While Not rst1.EOF
        mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_id,ybp_type,ybp_dn,ybp_tdn,ybp_money,ybp_num,ybp_oper,ybp_date)values ('" + rst1.Fields(0) + "','" + rst1.Fields(1) + "','" + rst1.Fields(2) + "','" + rst1.Fields(3) + "','" + rst1.Fields(4) + "','" + rst1.Fields(5) + "','" + rst1.Fields(6) + "','" + rst1.Fields(7) + "','" + rst1.Fields(8) + "')")
        SumMoney = SumMoney + rst1.Fields(5)
        rst1.MoveNext
        Wend
    '插入总计
    mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_money)values ('" + "总计" + "','" + Str(SumMoney) + "')")
    rst.Close
    rst1.Close
ElseIf Option2.Value Then       '打印指定日期内信息
SumMoney = 0
Dim date1 As String, date2 As String
date1 = Text2 & "-" & Combo1 & "-01"
date2 = Text3 & "-" & Combo2 & "-01"
    Set rst1 = mconn.Execute("select * from WTBDdb where  datediff('m',yb_date,'" + date1 + "')<=0 and datediff('m',yb_date,'" + date2 + "')>=0")
    If Not rst1.BOF Then rst1.MoveFirst
        While Not rst1.EOF
        mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_id,ybp_type,ybp_dn,ybp_tdn,ybp_money,ybp_num,ybp_oper,ybp_date)values ('" + rst1.Fields(0) + "','" + rst1.Fields(1) + "','" + rst1.Fields(2) + "','" + rst1.Fields(3) + "','" + rst1.Fields(4) + "','" + rst1.Fields(5) + "','" + rst1.Fields(6) + "','" + rst1.Fields(7) + "','" + rst1.Fields(8) + "')")
        SumMoney = SumMoney + rst1.Fields(5)
        rst1.MoveNext
        Wend
    '插入总计
   mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_money)values ('" + "总计" + "','" + Str(SumMoney) + "')")
    rst1.Close

ElseIf Option3.Value Then       '打印所有信息
SumMoney = 0
    Set rst1 = mconn.Execute("select * from WTBDdb ")
    If Not rst1.BOF Then rst1.MoveFirst
        While Not rst1.EOF
        mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_id,ybp_type,ybp_dn,ybp_tdn,ybp_money,ybp_num,ybp_oper,ybp_date)values ('" + rst1.Fields(0) + "','" + rst1.Fields(1) + "','" + rst1.Fields(2) + "','" + rst1.Fields(3) + "','" + rst1.Fields(4) + "','" + rst1.Fields(5) + "','" + rst1.Fields(6) + "','" + rst1.Fields(7) + "','" + rst1.Fields(8) + "')")
        SumMoney = SumMoney + rst1.Fields(5)
        rst1.MoveNext
        Wend
    '插入总计
    mconn.Execute ("insert into WTBDPdb(ybp_buyid,ybp_money)values ('" + "总计" + "','" + Str(SumMoney) + "')")
    rst1.Close
End If
'显示报表
Sleep 1000
DrpBuyDian.Show vbModal

Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub
