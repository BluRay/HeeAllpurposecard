VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBuyshuiP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户购水信息"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11580
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3960
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   0
      Width           =   735
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "FrmBuyshuiP.frx":0000
         Top             =   0
         Width           =   720
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "FrmBuyshuiP.frx":08AD
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   3
      FixedCols       =   2
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   495
      Left            =   4680
      OleObjectBlob   =   "FrmBuyshuiP.frx":0AE1
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmBuyshuiP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim YHno As String
Dim rst As Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path + "\B-Studio.skn"
Skin1.ApplySkin Me.hWnd
Dim dataitem As String

If YHModS Then
    If FrmYhMod.MSFlexGrid1.Enabled Then
        '显示详细信息
        FrmYhMod.MSFlexGrid1.Col = 0
        YHno = Trim$(FrmYhMod.MSFlexGrid1.Text)
        Set rst = mconn.Execute("select * from WTBdb where yb_id='" + YHno + "'")
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "没有任何信息！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "^购水编号|^用户编号|^购水日期            |^购水次数|^购水金额|^表一购买量|^表二购买量|^表三购买量|^表四购买量|^表一总购量|^表二总购量|^表三总购量|^表四总购量|"
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
    End If
Else

    If FrmYHcha.MSFlexGrid1.Enabled Then
        '显示详细信息
        FrmYHcha.MSFlexGrid1.Col = 0
        YHno = Trim$(FrmYHcha.MSFlexGrid1.Text)
        Set rst = mconn.Execute("select * from WTBdb where yb_id='" + YHno + "'")
    If Not rst.BOF Then rst.MoveFirst
    If rst.EOF Then
        MSFlexGrid1.Clear
        MSFlexGrid1.Enabled = False
        Beep
        MsgBox "没有任何信息！", vbOKOnly + vbInformation, App.Title
    Else
            With rst
            MSFlexGrid1.Rows = 1
            MSFlexGrid1.FormatString = "^购水编号|^用户编号|^购水日期            |^购水次数|^购水金额|^表一购买量|^表二购买量|^表三购买量|^表四购买量|^表一总购量|^表二总购量|^表三总购量|^表四总购量"
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

End If
Exit Sub
errhandle:
    MsgBox Error(ErR), vbOKOnly + vbInformation, App.Title

        
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
YHModS = False
End Sub
