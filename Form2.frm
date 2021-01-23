VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Maintain the Folder Transfer names"
   ClientHeight    =   4920
   ClientLeft      =   4530
   ClientTop       =   3570
   ClientWidth     =   6675
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   6675
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   3435
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   450
      Left            =   3480
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   450
      Left            =   2115
      TabIndex        =   4
      Top             =   4260
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   450
      Left            =   840
      TabIndex        =   3
      Top             =   4260
      Width           =   930
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   450
      Left            =   5250
      TabIndex        =   2
      Top             =   4260
      Width           =   930
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   165
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   2940
   End
   Begin VB.Label Label1 
      Caption         =   "Existing Folders"
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   210
      Width           =   1320
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim locNew

Private Sub cmdDelete_Click()
Dim f, hasData
' Check if TODOs exist

' match with array
For f = 0 To UBound(gSubArrays, 2)
    If UCase(gSubArrays(0, f)) = UCase(Text1.Text) Then  ' 0=name 1=array number
        Select Case f
        Case 0
            If gArray0(0) <> "" Then
                hasData = True
            End If
        Case 1
            If gArray1(0) <> "" Then
                hasData = True
            End If
        Case 2
            If gArray2(0) <> "" Then
                hasData = True
            End If
        Case 3
            If gArray3(0) <> "" Then
                hasData = True
            End If
        Case 4
            If gArray4(0) <> "" Then
                hasData = True
            End If
        Case 5
            If gArray5(0) <> "" Then
                hasData = True
            End If
        Case 6
            If gArray6(0) <> "" Then
                hasData = True
            End If
        Case 7
            If gArray7(0) <> "" Then
                hasData = True
            End If
        Case 8
            If gArray8(0) <> "" Then
                hasData = True
            End If
        Case 9
            If gArray9(0) <> "" Then
                hasData = True
            End If
        End Select
        Exit For
    End If
Next f

If hasData Then
    MsgBox "That folder still has data"
Else
    List1.RemoveItem List1.ListIndex
End If

Text1.Visible = False

cmdUpdate.Visible = False
cmdDelete.Visible = False
cmdNew.Visible = True

frmMain.cmdSave.Visible = True

End Sub

Private Sub cmdExit_Click()

Dim FileNum, fileRec, f
On Error Resume Next

FileNum = FreeFile
Open gAppPath & "folders.dat" For Output As FileNum
For f = 0 To List1.ListCount - 1
    List1.ListIndex = f
    Print #FileNum, List1.Text
Next f
Close FileNum

Unload Me
End Sub

Private Sub cmdNew_Click()

If UBound(gSubArrays, 2) = 10 Then
    MsgBox "All 10 folders in use"
    Exit Sub
End If

Text1.Visible = True
Text1.Text = ""

cmdNew.Visible = False
cmdUpdate.Visible = True

locNew = True

End Sub

Private Sub cmdUpdate_Click()

Dim f, matched


If locNew = False Then
    ' UPDATE - Change name
    For f = 0 To UBound(gSubArrays, 2) - 1
        If UCase(gSubArrays(0, f)) = UCase(List1.Text) Then
            gSubArrays(0, f) = Text1.Text
            frmMain.cmdFolder(gSubArrays(1, f)).Caption = Text1.Text
            Exit For
        End If
    Next f
    
    If Trim(Text1.Text) <> "" Then
        List1.RemoveItem List1.ListIndex
        List1.AddItem Text1.Text
    End If
Else
    ' NEW, Check unique.
    matched = False
    For f = 0 To UBound(gSubArrays, 2) - 1
        If UCase(gSubArrays(0, f)) = UCase(List1.Text) Then
            matched = True
            Exit For
        End If
    Next f
    If Not matched Then
        If List1.ListCount <> 0 Then
            Load frmMain.cmdFolder(f)
            frmMain.cmdFolder(f).Left = frmMain.cmdFolder(0).Left
            frmMain.cmdFolder(f).Top = frmMain.cmdFolder(f - 1).Top + frmMain.cmdFolder(0).Height
        End If
        frmMain.cmdFolder(f).Caption = Text1.Text
        frmMain.cmdFolder(f).Visible = True
        gSubArrays(0, f) = Text1.Text  ' name
        gSubArrays(1, f) = f     ' array number
        ReDim Preserve gSubArrays(1, UBound(gSubArrays, 2) + 1)
        List1.AddItem Text1.Text
    Else
        MsgBox "That folder name already exists"
        Text1.SetFocus
    End If
    
End If

Text1.Visible = False

cmdUpdate.Visible = False
cmdDelete.Visible = False
cmdNew.Visible = True

End Sub

Private Sub Form_Load()
Dim FileNum, fileRec
On Error Resume Next
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

FileNum = FreeFile
Open gAppPath & "folders.dat" For Input As FileNum
If Err = 0 Then
    Do While Not EOF(FileNum)
        Input #FileNum, fileRec
        List1.AddItem fileRec
    Loop
End If
Close FileNum

End Sub


Private Sub List1_Click()
 Text1.Visible = True
Text1.Text = List1.Text

cmdNew.Visible = False
cmdUpdate.Visible = True
cmdDelete.Visible = True

locNew = False

End Sub


