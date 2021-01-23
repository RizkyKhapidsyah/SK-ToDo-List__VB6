VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   Caption         =   "Picture Library -  Print"
   ClientHeight    =   4530
   ClientLeft      =   6510
   ClientTop       =   4770
   ClientWidth     =   4155
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4155
   Begin MSComDlg.CommonDialog SelectPrinter 
      Left            =   1785
      Top             =   1965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Spacing"
      ForeColor       =   &H00C00000&
      Height          =   1365
      Left            =   2145
      TabIndex        =   9
      Top             =   1590
      Width           =   1320
      Begin VB.OptionButton optS 
         Caption         =   "Single"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   360
         Width           =   1110
      End
      Begin VB.OptionButton optS 
         Caption         =   "Double"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   765
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Size"
      ForeColor       =   &H00C00000&
      Height          =   1365
      Left            =   600
      TabIndex        =   5
      Top             =   1590
      Width           =   1320
      Begin VB.OptionButton optF 
         Caption         =   "8"
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optF 
         Caption         =   "10"
         Height          =   315
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
      Begin VB.OptionButton optF 
         Caption         =   "12"
         Height          =   315
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   975
         Width           =   1035
      End
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2722
      TabIndex        =   4
      Top             =   3885
      Width           =   1170
   End
   Begin VB.CommandButton bOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   1462
      TabIndex        =   3
      Top             =   3885
      Width           =   1170
   End
   Begin VB.CommandButton bPrinter 
      Caption         =   "&Printer Setup"
      Height          =   435
      Left            =   172
      TabIndex        =   2
      Top             =   3885
      Width           =   1170
   End
   Begin VB.TextBox tCopies 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Left            =   1695
      TabIndex        =   1
      Text            =   "1"
      Top             =   555
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Number of copies:"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   1530
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bCancel_Click()
'D Ind cancelled
gPCancel = True
Unload Me
End Sub

Private Sub bOK_Click()

'D Globalise number of copies
gNumCopies = tCopies.Text

Unload Me
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdSetup_Click()
    SelectPrinter.Flags = &HC0004
    '40000 = returns '# copies'
    '80000 = disable 'to file'
    '04    = disable Selection
    SelectPrinter.Action = 5
    'D If copies not set
    If tCopies.Text = 1 Then
        'D Take from dialog
        tCopies.Text = SelectPrinter.copies
    End If

End Sub

Private Sub Form_Load()

'Move Screen.Width / 2 - Width / 2, Screen.Height / 2 - Height / 2
Call CentreForm(Me)
gPCancel = False
'D If font not previously set
If gFontSize = "" Then
    gFontSize = 10
    optF(1).Value = True
'D Else
Else
    'D Reset per previous selection
    Select Case gFontSize
        Case 8
        optF(0).Value = True
        Case 10
        optF(1).Value = True
    End Select
End If

'D If spacing not previously set
If gSpacing = "" Then
    'D Default to single
    gSpacing = 1
    optS(0).Value = True
'D Else
Else
    'D Reset per previous selection
    Select Case gSpacing
        Case 1
        optS(0).Value = True
        Case 2
        optS(1).Value = True
    End Select
'    optS(0).Value = True
End If
    


End Sub

Private Sub optF_Click(Index As Integer)
'D Set font size
Select Case Index
    Case 0
    gFontSize = 8
    Case 1
    gFontSize = 10
    Case 2
    gFontSize = 12
End Select
End Sub


Private Sub optS_Click(Index As Integer)
Select Case Index
    Case 0
    gSpacing = 1
    Case 1
    gSpacing = 2
End Select
End Sub


