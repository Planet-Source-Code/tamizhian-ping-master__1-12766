VERSION 5.00
Begin VB.Form frmPingMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ping Master by Tamizhian@hotmail.com"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   3195
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1500
      Width           =   3315
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "Find All"
      Height          =   375
      Left            =   2220
      TabIndex        =   9
      Top             =   1020
      Width           =   1155
   End
   Begin VB.CommandButton cmdGetName 
      Caption         =   "Get Name"
      Height          =   375
      Left            =   2220
      TabIndex        =   8
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton cmdGetIP 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   2220
      TabIndex        =   7
      Top             =   60
      Width           =   1155
   End
   Begin VB.TextBox txtIpTo 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   6
      Top             =   1020
      Width           =   435
   End
   Begin VB.TextBox txtIpTo 
      Height          =   375
      Index           =   2
      Left            =   1140
      TabIndex        =   5
      Top             =   1020
      Width           =   435
   End
   Begin VB.TextBox txtIpFrom 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox txtIpFrom 
      Height          =   375
      Index           =   2
      Left            =   1140
      TabIndex        =   3
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox txtIpFrom 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox txtIpFrom 
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox txtIPName 
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frmPingMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFindAll_Click()
Static YesExit As Boolean

If cmdFindAll.Caption <> "Stop" Then
    cmdFindAll.Caption = "Stop"
    txtLog.Text = ""
    Me.MousePointer = vbHourglass
    On Error GoTo OutOfHere
    For i = txtIpFrom(2) To txtIpTo(2)
        For j = txtIpFrom(3) To txtIpTo(3)
            mip = txtIpFrom(0) & "." & txtIpFrom(1) & "." & i & "." & j
            mname = IpToAddr(mip)
            
            txtLog.Text = txtLog.Text & mip & "   " & mname & vbCrLf
            If YesExit Then YesExit = False: Exit For
            DoEvents
        Next
    Next
OutOfHere:
    cmdFindAll.Caption = "Find All"
    Me.MousePointer = vbNormal
Else
    cmdFindAll.Caption = "Find All"
    YesExit = True
End If
End Sub

Private Sub cmdGetIP_Click()
    For i = 0 To 3: txtIpFrom(i) = "": Next
    Me.MousePointer = vbHourglass
    
    mip = AddrToIP(txtIPName)
    
    mpos1 = InStr(1, mip, ".")
    txtIpFrom(0) = Left(mip, mpos1 - 1)
    mpos2 = InStr(mpos1 + 1, mip, ".")
    txtIpFrom(1) = Mid(mip, mpos1 + 1, mpos2 - (mpos1 + 1))
    mpos3 = InStr(mpos2 + 1, mip, ".")
    txtIpFrom(2) = Mid(mip, mpos2 + 1, mpos3 - (mpos2 + 1))
    txtIpFrom(3) = Mid(mip, mpos3 + 1)
    
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdGetName_Click()
    txtIPName = ""
    
    Me.MousePointer = vbHourglass
    mip = txtIpFrom(0) & "." & txtIpFrom(1) & "." & txtIpFrom(2) & "." & txtIpFrom(3)
    txtIPName = IpToAddr(mip)
    Me.MousePointer = vbNormal
    
End Sub

Private Sub Form_Load()
    StartWinsock "ABC"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    EndWinsock
End Sub
