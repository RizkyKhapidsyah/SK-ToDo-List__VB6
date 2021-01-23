VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Print Preview"
   ClientHeight    =   8595
   ClientLeft      =   1305
   ClientTop       =   1800
   ClientWidth     =   11880
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.Frame fraPrinter 
      Caption         =   "Printer settings "
      Height          =   5115
      Left            =   7170
      TabIndex        =   22
      Top             =   1410
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdSetup 
         Caption         =   "&Printer Setup"
         Height          =   435
         Left            =   345
         TabIndex        =   16
         Top             =   4335
         Width           =   1170
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   3660
         TabIndex        =   18
         Top             =   4335
         Width           =   885
      End
      Begin VB.CommandButton cmdCancelP 
         Caption         =   "C&ancel"
         Height          =   375
         Index           =   1
         Left            =   2250
         TabIndex        =   17
         Top             =   4335
         Width           =   885
      End
      Begin VB.Frame Frame5 
         Caption         =   "Spacing"
         Height          =   1425
         Left            =   2520
         TabIndex        =   27
         Top             =   1770
         Width           =   2055
         Begin VB.OptionButton Option8 
            Caption         =   "Single"
            Height          =   315
            Left            =   210
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Double"
            Height          =   285
            Left            =   210
            TabIndex        =   14
            Top             =   660
            Width           =   1635
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Font "
         Height          =   1455
         Left            =   120
         TabIndex        =   26
         Top             =   1770
         Width           =   2070
         Begin VB.OptionButton Option10 
            Caption         =   "Courier New"
            Height          =   285
            Left            =   210
            TabIndex        =   9
            Top             =   1020
            Width           =   1635
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Arial"
            Height          =   285
            Left            =   210
            TabIndex        =   8
            Top             =   660
            Width           =   1635
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Times New Roman"
            Height          =   315
            Left            =   210
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   1755
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2670
         TabIndex        =   15
         Text            =   "1"
         Top             =   3600
         Width           =   645
      End
      Begin VB.Frame Frame3 
         Caption         =   "Font Size "
         Height          =   1395
         Left            =   2520
         TabIndex        =   24
         Top             =   300
         Width           =   2025
         Begin VB.OptionButton Option5 
            Caption         =   "12"
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   990
            Width           =   1365
         End
         Begin VB.OptionButton Option4 
            Caption         =   "10"
            Height          =   285
            Left            =   240
            TabIndex        =   11
            Top             =   660
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.OptionButton Option3 
            Caption         =   "8"
            Height          =   345
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Paper "
         Height          =   1395
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "Portrait"
            Height          =   375
            Left            =   210
            TabIndex        =   5
            Top             =   420
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Landscape"
            Height          =   375
            Left            =   210
            TabIndex        =   6
            Top             =   750
            Width           =   1515
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Copies:"
         Height          =   315
         Left            =   1110
         TabIndex        =   25
         Top             =   3660
         Width           =   1425
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   525
      Top             =   5505
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDisk 
      Caption         =   "&Disk"
      Height          =   345
      Left            =   2070
      TabIndex        =   1
      Top             =   8070
      Width           =   885
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   345
      Left            =   4425
      TabIndex        =   2
      Top             =   8070
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   4155
      TabIndex        =   20
      Top             =   2715
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   1380
         TabIndex        =   19
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Printing ....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   1140
         TabIndex        =   21
         Top             =   150
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   8925
      TabIndex        =   4
      Top             =   8070
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   6675
      TabIndex        =   3
      Top             =   8070
      Width           =   885
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   11835
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim locFilePath

Private Sub cmdCancel_Click()
gCancelPrinting = True
Screen.MousePointer = 0
cmdCancel.Visible = False
 

End Sub

Private Sub cmdCancelP_Click(Index As Integer)

fraPrinter.Visible = False

End Sub

Private Sub cmdCopy_Click()
Dim f, CopyText

lblPrompt.Caption = "Copying ....."
Frame1.Visible = True

If List1.ListCount = -1 Then
    Exit Sub
Else
    'Printer.FontName = "Times New Roman"
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    For f = 0 To List1.ListCount - 1
        DoEvents
        If gCancelPrinting Then
            gCancelPrinting = False
            Frame1.Visible = False
            Exit For
        End If
        List1.ListIndex = f
        CopyText = CopyText & List1.Text & Chr(13) & Chr(10)
    Next f
End If
Clipboard.Clear
Clipboard.SetText CopyText
MsgBox "Copied to clipboard for e.g. paste into Notepad"
Frame1.Visible = False
List1.ListIndex = 0
End Sub


Private Sub cmdDisk_Click()
Dim f
On Error Resume Next

If List1.ListCount = -1 Then
    Exit Sub
Else

    ' Get the CV file location
    
    CommonDialog1.InitDir = gAppPath
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All text files (*.TXT)|*.TXT"
    CommonDialog1.DefaultExt = "*.*"
    CommonDialog1.FileName = "RecPrint.TXT"  'set sample name
    CommonDialog1.DialogTitle = "The file to be created"
    CommonDialog1.Action = 1
    
    'D If Cancelled
    If Err <> 0 Then
        'D No processing
        Exit Sub
    'D Else
    Else
        'D Save path
        locFilePath = CommonDialog1.FileName
    End If
    
    If Dir(locFilePath) <> "" Then
        If MsgBox("File already exists. OK to overwrite?", 20) <> vbYes Then
            Exit Sub
        End If
    End If
    
    Frame1.Visible = True

    Open locFilePath For Output As #5
    For f = 0 To List1.ListCount - 1
        DoEvents
        If gCancelPrinting Then
            gCancelPrinting = False
            Frame1.Visible = False
            Exit For
        End If
        List1.ListIndex = f
        Print #5, List1.Text
    Next f
        
    Frame1.Visible = False
    MsgBox "The report has been output to file " & locFilePath
    Close 5
End If
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim f, g, copies

If Option1 Then
    Printer.Orientation = vbPRORPortrait
End If
If Option2 Then
    Printer.Orientation = vbPRORLandscape
End If
If Option3 Then
    Printer.FontSize = 8
End If
If Option4 Then
    Printer.FontSize = 10
End If
If Option5 Then
    Printer.FontSize = 12
End If
If Option6 Then
    Printer.FontName = "Times New Roman"
End If
If Option7 Then
    Printer.FontName = "Arial"
End If
If Option10 Then
    Printer.FontName = "Courier New"
End If
' Option8 = single spacing, Option9 = double

copies = 1
If IsNumeric(Text1.Text) Then
    copies = Val(Text1.Text)
End If

fraPrinter.Visible = True

lblPrompt.Caption = "Printing ....."
Frame1.Visible = True

If List1.ListCount = -1 Then
    Exit Sub
Else
    For g = 1 To copies
        If g > 1 Then
            Printer.NewPage
        End If
        For f = 0 To List1.ListCount - 1
           DoEvents
           If gCancelPrinting Then
               gCancelPrinting = False
               Frame1.Visible = False
               Exit For
           End If
           List1.ListIndex = f
           Printer.Print List1.Text
           If Option9 Then
                Printer.Print " "
           End If
        Next f
    Next g
    
End If
Printer.EndDoc
Unload Me

End Sub

Private Sub cmdPrint_Click()

fraPrinter.Visible = True

End Sub


Private Sub cmdSetup_Click()
    CommonDialog1.Flags = &HC0004
    '40000 = returns '# copies'
    '80000 = disable 'to file'
    '04    = disable Selection
    CommonDialog1.Action = 5
    'D If copies not set
    If Text1.Text = 1 Then
        'D Take from dialog
        Text1.Text = CommonDialog1.copies
    End If
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 
fraPrinter.Move 4020, 2220
' TRIED STARTING AT MAX SCREEN SIZE BUT
' PRINTING IS A COPY OF THE SCREEN!!

'List1.Width = Screen.Width
'List1.Height = Screen.Height - (cmdExit.Height * 3)
'cmdDisk.Top = Screen.Height - (cmdExit.Height * 2.5)
'cmdCopy.Top = cmdDisk.Top
'cmdPrint.Top = cmdDisk.Top
'cmdExit.Top = cmdDisk.Top


'''' WON'T PRINT
'If Screen.Width Then
'    gLengthOfLine = 90
'End If
End Sub


