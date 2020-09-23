VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extract File(s)"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "&Browse..."
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   2010
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.CheckBox chk 
      Caption         =   "Overwrite w/o asking"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox chk 
      Caption         =   "Space to Underscore"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CheckBox chk 
      Caption         =   "Extract only Newer"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox chk 
      Caption         =   "Ignore Folder Names"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Extract"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton opt 
      Caption         =   "Selected Files"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton opt 
      Caption         =   "All Files"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Password:"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Extract To:"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim ltInfo As tUnzipInfo
            With ltInfo
                .ExtractToPath = txt(0).Text
                If opt(1).Value Then frmZipper.GetSelFiles .Include
                If txt(1).Enabled Then .Password = txt(1).Text
                If chk(0).Value = vbChecked Then .Attributes = zipDisregardFolderNames
                If chk(1).Value = vbChecked Then .Attributes = .Attributes Or zipExtractOnlyNewer
                If chk(2).Value = vbChecked Then .Attributes = .Attributes Or zipSpaceToUnderscore
                If chk(3).Value = vbChecked Then .OverwriteAll = True
            End With
            Hide
            frmZipper.UnzipFiles ltInfo
        Case 1
            Hide
        Case 2
            Dim lsTemp As String
            lsTemp = BrowseForFolder(hWnd, "Choose a folder for extraction.", vbNullString)
            If LenB(lsTemp) > 0 Then txt(0).Text = lsTemp
    End Select
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If frmZipper.ListCount(True) > 0 Then
        opt(1).Enabled = True
        opt(1).Value = True
    Else
        opt(0).Value = True
        opt(1).Enabled = False
    End If
    txt(1).Enabled = frmZipper.Encrypted
    txt(0).SetFocus
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 0 Then
        Dim lsTemp As String
        lsTemp = txt(Index).Text
        If Right$(lsTemp, 1) = "\" Then lsTemp = Left$(lsTemp, Len(lsTemp) - 1)
        cmd(0).Enabled = FolderExists(PathGetParentFolder(lsTemp)) And FileNameIsLegal(PathGetFileName(lsTemp))
        ac_Change txt(Index), txt(Index).Tag, acbFolder
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then txt(0).Tag = KeyCode
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    ac_KeyPress txt(Index), KeyAscii, acbFolder
End Sub
