VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MsysSetD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "电表参数设置"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9360
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8640
      OleObjectBlob   =   "MsysSetD.frx":0000
      Top             =   5520
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3480
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "MsysSetD.frx":0234
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   495
      Left            =   4200
      OleObjectBlob   =   "MsysSetD.frx":0AE1
      TabIndex        =   17
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   9015
      Begin VB.Frame Frame5 
         Caption         =   "添加用电类型："
         Height          =   1935
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   8775
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   6120
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   6120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            Caption         =   "添 加"
            Height          =   375
            Left            =   1800
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "保 存"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5520
            TabIndex        =   5
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   6120
            TabIndex        =   10
            Top             =   960
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   5160
            OleObjectBlob   =   "MsysSetD.frx":0B42
            TabIndex        =   11
            Top             =   1080
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "MsysSetD.frx":0BAA
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "MsysSetD.frx":0C18
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   4800
            OleObjectBlob   =   "MsysSetD.frx":0C80
            TabIndex        =   15
            Top             =   720
            Width           =   3735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "MsysSetD.frx":0D46
            TabIndex        =   16
            Top             =   720
            Width           =   3375
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   2280
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   5
         FormatString    =   "^用电类型      |^单价(元)    |^允许过载次数    |^允许透支额度    |^设置日期                     "
      End
   End
End
Attribute VB_Name = "MsysSetD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As Recordset
Dim m_QuerySQLstr As String

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text2.Enabled = True
Text6.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text7.Enabled = True
Text2 = ""
Text6 = ""
Text3 = ""
Text4 = ""
    Text7.Text = Str(Year(Now)) & "-" & FormatString(Str(Month(Now)), 2) & "-" & FormatString(Str(Day(Now)), 2)
Text2.SetFocus
Command4.Enabled = True
End Sub

Private Sub Command4_Click()
On Error GoTo errhandle

If Command3.Enabled Then
If Text2 = "" Or Text6 = "" Then
MsgBox "所有项目为必填项，检查是否有漏填！"
Exit Sub
End If
mconn.Execute ("insert into WTDdb (Ds_name,Ds_price,Ds_gznum,Ds_tz,Ds_Date)values('" + Text2 + "','" + Text6 + "','" + Text4 + "','" + Text3 + "','" + Text7 + "')")
MsgBox "添加成功！"
    Call RefreshDianType(m_QuerySQLstr)
Command4.Enabled = False
Text2.Enabled = 0
Text6.Enabled = 0
Text3.Enabled = 0
Text4.Enabled = 0
Text7.Enabled = 0
Else        '当添加按钮不可用时 ，则为修改数据
If Text2 = "" Or Text6 = "" Then
MsgBox "所有项目为必填项，检查是否有漏填！"
Exit Sub
End If
mconn.Execute ("update WTDdb set Ds_price='" + Text6 + "',DS_gznum='" + Text4 + "',Ds_tz='" + Text3 + "',Ds_date='" + Text7 + "' where Ds_name='" + Text2 + "' ")
MsgBox "修改成功！"
    Call RefreshDianType(m_QuerySQLstr)
Command4.Enabled = False


End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

End Sub

Private Sub Form_Load()
 Skin1.LoadSkin App.Path + "\B-Studio.skn"
 Skin1.ApplySkin Me.hWnd
    m_QuerySQLstr = "select * from WTDdb"
    Call RefreshDianType(m_QuerySQLstr)

End Sub

Sub RefreshDianType(m_QuerySQLString As String)   '显示所有用电类型
On Error GoTo errhandle
Dim dataitem As String
Set rst = mconn.Execute(m_QuerySQLString)
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "现在没有任何信息！请添加！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            dataitem = ""
            While Not rst.EOF
                    dataitem = .Fields("Ds_name") + vbTab
                    dataitem = dataitem + Str(.Fields("Ds_price")) + vbTab
                    dataitem = dataitem + Str(.Fields("Ds_gznum")) + vbTab
                    dataitem = dataitem + Str(.Fields("Ds_tz")) + vbTab
                    dataitem = dataitem + .Fields("Ds_Date") + vbTab
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

Private Sub MSFlexGrid1_Click()
'修改现有数据
'填充此项数据
MSFlexGrid1.Col = 0
Text2 = MSFlexGrid1.Text
MSFlexGrid1.Col = 1
Text6 = MSFlexGrid1.Text
MSFlexGrid1.Col = 2
Text4 = MSFlexGrid1.Text
MSFlexGrid1.Col = 3
Text3 = MSFlexGrid1.Text
MSFlexGrid1.Col = 4
Text7 = MSFlexGrid1.Text
Command3.Enabled = False
Command4.Enabled = True
Text6.Enabled = True
Text4.Enabled = True
Text3.Enabled = True
    Text7.Text = Str(Year(Now)) & "-" & FormatString(Str(Month(Now)), 2) & "-" & FormatString(Str(Day(Now)), 2)
End Sub
