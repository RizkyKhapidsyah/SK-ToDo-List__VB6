VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   " MY TODO LIST - IN PRIORITY ORDER Ver 2"
   ClientHeight    =   8595
   ClientLeft      =   885
   ClientTop       =   1800
   ClientWidth     =   11865
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11865
   Begin VB.CommandButton cmdMaintain 
      Caption         =   "&Maintain"
      Height          =   450
      Left            =   10365
      TabIndex        =   12
      Top             =   7500
      Width           =   1500
   End
   Begin VB.CommandButton cmdCompleted 
      Caption         =   "&Completed"
      Height          =   450
      Left            =   10365
      TabIndex        =   11
      Top             =   7020
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   4500
      TabIndex        =   4
      Top             =   8070
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   450
      Left            =   10680
      TabIndex        =   7
      Top             =   8070
      Width           =   930
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   450
      Left            =   9435
      TabIndex        =   6
      Top             =   8070
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   450
      Left            =   8145
      TabIndex        =   5
      Top             =   8070
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdAmend 
      Caption         =   "A&mend"
      Height          =   450
      Left            =   3405
      TabIndex        =   3
      Top             =   8070
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdAddStart 
      Caption         =   "Add to &Start"
      Height          =   270
      Left            =   2235
      TabIndex        =   2
      Top             =   8325
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdAddEnd 
      Caption         =   "Add to &End"
      Height          =   270
      Left            =   2220
      TabIndex        =   1
      Top             =   8010
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   450
      Left            =   1215
      TabIndex        =   0
      Top             =   8070
      Width           =   930
   End
   Begin VB.TextBox txtEntry 
      Height          =   360
      Left            =   195
      TabIndex        =   15
      Top             =   7635
      Visible         =   0   'False
      Width           =   9795
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   6810
      Left            =   9990
      TabIndex        =   14
      Top             =   435
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   12012
      _Version        =   393216
      Appearance      =   0
      LargeChange     =   5
      Orientation     =   1638400
   End
   Begin VB.CommandButton cmdFolder 
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1170
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdALL 
      Caption         =   "&ALL"
      Height          =   450
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   690
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "&General"
      Height          =   450
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   7020
      Left            =   45
      TabIndex        =   20
      Top             =   270
      Width           =   10260
      Begin VB.Label lblToDo 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   660
         TabIndex        =   22
         ToolTipText     =   "Drag to move, Right click to change"
         Top             =   180
         Width           =   9255
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   630
         X2              =   2340
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lblNumb 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   0
         Left            =   75
         TabIndex        =   21
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Image imgMoveIt 
      Height          =   480
      Left            =   7170
      Picture         =   "Form1.frx":08CA
      Top             =   8235
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblCurrentFolder 
      Caption         =   "FOLDER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1890
      TabIndex        =   19
      Top             =   60
      Width           =   5445
   End
   Begin VB.Label label1 
      Caption         =   "CURRENT FOLDER:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   195
      TabIndex        =   18
      Top             =   60
      Width           =   1620
   End
   Begin VB.Label lblActionMsg 
      Caption         =   "Action messages"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3555
      TabIndex        =   17
      Top             =   7410
      Width           =   6375
   End
   Begin VB.Image imgBin 
      Height          =   585
      Left            =   150
      Picture         =   "Form1.frx":1194
      Top             =   7995
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Right click an entry to amend it"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   7425
      Width           =   2295
   End
   Begin VB.Label label1 
      Caption         =   "FOLDERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   0
      Left            =   10650
      TabIndex        =   13
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAddEnd_Click()
Dim f

If Trim(txtEntry.Text) = "" Then
    Exit Sub
End If

Select Case gCurrentFolder
Case -1
    gMainArray(UBound(gMainArray)) = txtEntry.Text
    ReDim Preserve gMainArray(UBound(gMainArray) + 1)
Case 0
    gArray0(UBound(gArray0)) = txtEntry.Text
    ReDim Preserve gArray0(UBound(gArray0) + 1)
Case 1
    gArray1(UBound(gArray1)) = txtEntry.Text
    ReDim Preserve gArray1(UBound(gArray1) + 1)
Case 2
    gArray2(UBound(gArray2)) = txtEntry.Text
    ReDim Preserve gArray2(UBound(gArray2) + 1)
Case 3
    gArray3(UBound(gArray3)) = txtEntry.Text
    ReDim Preserve gArray3(UBound(gArray3) + 1)
Case 4
    gArray4(UBound(gArray4)) = txtEntry.Text
    ReDim Preserve gArray4(UBound(gArray4) + 1)
Case 5
    gArray5(UBound(gArray5)) = txtEntry.Text
    ReDim Preserve gArray5(UBound(gArray5) + 1)
Case 6
    gArray6(UBound(gArray6)) = txtEntry.Text
    ReDim Preserve gArray6(UBound(gArray6) + 1)
Case 7
    gArray7(UBound(gArray7)) = txtEntry.Text
    ReDim Preserve gArray7(UBound(gArray7) + 1)
Case 8
    gArray8(UBound(gArray8)) = txtEntry.Text
    ReDim Preserve gArray8(UBound(gArray8) + 1)
Case 9
    gArray9(UBound(gArray9)) = txtEntry.Text
    ReDim Preserve gArray9(UBound(gArray9) + 1)
End Select

gNumEntries = gNumEntries + 1

' Move display to last page

If gNumEntries <= gMaxLinesOnScreen Then
    Call FlatScrollBar1_Change
    FlatScrollBar1.Value = 0
Else
    FlatScrollBar1.Visible = True
    FlatScrollBar1.Max = gMaxLinesOnScreen
    FlatScrollBar1.Value = gNumEntries - (gMaxLinesOnScreen - 1)
End If

cmdCancel.Visible = False
txtEntry.Text = ""
txtEntry.Visible = False
cmdNew.Visible = True
cmdAmend.Visible = False
cmdAddEnd.Visible = False
cmdAddStart.Visible = False

cmdNew.SetFocus
End Sub

Private Sub cmdAddStart_Click()
Dim f

If Trim(txtEntry.Text) = "" Then
    Exit Sub
End If

Select Case gCurrentFolder
Case -1
    ' Move array entries down one
    For f = UBound(gMainArray) To 1 Step -1
        gMainArray(f) = gMainArray(f - 1)
    Next f
    ' Slot in new entry
    gMainArray(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gMainArray(UBound(gMainArray) + 1)

Case 0
    ' Move array entries down one
    For f = UBound(gArray0) To 1 Step -1
        gArray0(f) = gArray0(f - 1)
    Next f
    ' Slot in new entry
    gArray0(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray0(UBound(gArray0) + 1)
Case 1
    ' Move array entries down one
    For f = UBound(gArray1) To 1 Step -1
        gArray1(f) = gArray1(f - 1)
    Next f
    ' Slot in new entry
    gArray1(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray1(UBound(gArray1) + 1)
Case 2
    ' Move array entries down one
    For f = UBound(gArray2) To 1 Step -1
        gArray2(f) = gArray2(f - 1)
    Next f
    ' Slot in new entry
    gArray2(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray2(UBound(gArray2) + 1)
Case 3
    ' Move array entries down one
    For f = UBound(gArray3) To 1 Step -1
        gArray3(f) = gArray3(f - 1)
    Next f
    ' Slot in new entry
    gArray3(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray3(UBound(gArray3) + 1)
Case 4
    ' Move array entries down one
    For f = UBound(gArray4) To 1 Step -1
        gArray4(f) = gArray4(f - 1)
    Next f
    ' Slot in new entry
    gArray4(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray4(UBound(gArray4) + 1)
Case 5
    ' Move array entries down one
    For f = UBound(gArray5) To 1 Step -1
        gArray5(f) = gArray5(f - 1)
    Next f
    ' Slot in new entry
    gArray5(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray5(UBound(gArray5) + 1)
Case 6
    ' Move array entries down one
    For f = UBound(gArray6) To 1 Step -1
        gArray6(f) = gArray6(f - 1)
    Next f
    ' Slot in new entry
    gArray6(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray6(UBound(gArray6) + 1)
Case 7
    ' Move array entries down one
    For f = UBound(gArray7) To 1 Step -1
        gArray7(f) = gArray7(f - 1)
    Next f
    ' Slot in new entry
    gArray7(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray7(UBound(gArray7) + 1)
Case 8
    ' Move array entries down one
    For f = UBound(gArray8) To 1 Step -1
        gArray8(f) = gArray8(f - 1)
    Next f
    ' Slot in new entry
    gArray8(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray8(UBound(gArray8) + 1)
Case 9
    ' Move array entries down one
    For f = UBound(gArray9) To 1 Step -1
        gArray9(f) = gArray9(f - 1)
    Next f
    ' Slot in new entry
    gArray9(0) = txtEntry.Text
    ' Increase array size
    ReDim Preserve gArray9(UBound(gArray9) + 1)
End Select

' Display
'(IF 1st load)
gNumEntries = gNumEntries + 1
If FlatScrollBar1.Value = 0 Then
    Call FlatScrollBar1_Change
Else
  FlatScrollBar1.Value = 0
End If

cmdCancel.Visible = False
txtEntry.Text = ""
txtEntry.Visible = False
cmdNew.Visible = True
cmdAmend.Visible = False
cmdAddEnd.Visible = False
cmdAddStart.Visible = False

cmdNew.SetFocus
End Sub


Private Sub cmdALL_Click()
' Set FOLDER name
lblCurrentFolder.Caption = "ALL"
gCurrentFolder = 99

imgBin.Visible = False

End Sub

Private Sub cmdAmend_Click()
Dim f

lblToDo(gSelIndex).Caption = txtEntry.Text

f = Val(gCurrentFolder)
Select Case f
Case -1
    gMainArray(gSelIndex + gCurrent1) = txtEntry.Text
Case 0
    gArray0(gSelIndex + gCurrent1) = txtEntry.Text
Case 1
    gArray1(gSelIndex + gCurrent1) = txtEntry.Text
Case 2
    gArray2(gSelIndex + gCurrent1) = txtEntry.Text
Case 3
    gArray3(gSelIndex + gCurrent1) = txtEntry.Text
Case 4
    gArray4(gSelIndex + gCurrent1) = txtEntry.Text
Case 5
    gArray5(gSelIndex + gCurrent1) = txtEntry.Text
Case 6
    gArray6(gSelIndex + gCurrent1) = txtEntry.Text
Case 7
    gArray7(gSelIndex + gCurrent1) = txtEntry.Text
Case 8
    gArray8(gSelIndex + gCurrent1) = txtEntry.Text
Case 9
    gArray9(gSelIndex + gCurrent1) = txtEntry.Text
End Select

cmdCancel.Visible = False
txtEntry.Visible = False
txtEntry.Text = ""
cmdNew.Visible = True
cmdAmend.Visible = False
 
cmdSave.Visible = True
cmdNew.SetFocus


End Sub

Private Sub cmdCancel_Click()

    txtEntry.Text = ""
    txtEntry.Visible = False
    cmdAmend.Visible = False
    cmdNew.Visible = True
    cmdAddEnd.Visible = False
    cmdAddStart.Visible = False
    
    gSelIndex = ""
    
    txtEntry.Visible = False
    cmdCancel.Visible = False

End Sub

Private Sub cmdCompleted_Click()
Dim f

' IGNORE IF SAME FOLDER
If lblCurrentFolder.Caption = Right(cmdCompleted.Caption, Len(cmdCompleted.Caption) - 1) Then
    Exit Sub
End If

cmdNew.Visible = False
gCurrent1 = 0

' Set FOLDER name
lblCurrentFolder.Caption = "THE COMPLETED LIST"
gCurrentFolder = 99

' unload current display
' For each entry
For f = 1 To gNumberShowing
        ' unLoad lines
        Unload Line1(f)
        ' unLoad Todo boxes
        Unload lblToDo(f)
        ' unLoad Todo numbers
        Unload lblNumb(f)
Next f
gNumberShowing = 0
gNumEntries = UBound(gCompletedArray)
lblToDo(0).Caption = gCompletedArray(0)

lblNumb(0).Caption = "1"

' For each entry
For f = 1 To gNumEntries - 1
    ' If not exceeded screen display
    If f <= gMaxLinesOnScreen Then
        ' Load lines
        Load Line1(f)
        Line1(f).X1 = Line1(0).X1  'left
        Line1(f).X2 = Line1(0).X2  ' right
        Line1(f).Y1 = Line1(f - 1).Y1 + lblToDo(0).Height + 20
        Line1(f).Y2 = Line1(f).Y1
        Line1(f).Visible = True
        ' Load Todo boxes
        Load lblToDo(f)
        lblToDo(f).Top = lblToDo(f - 1).Top + lblToDo(0).Height + 20
        lblToDo(f).Caption = gCompletedArray(f)
        lblToDo(f).Visible = True
        gNumberShowing = gNumberShowing + 1 ' total showing
        ' Load Todo numbers
        Load lblNumb(f)
        lblNumb(f).Top = lblToDo(f).Top
        lblNumb(f).Caption = lblNumb(f - 1).Caption + 1
        
        lblNumb(f).Visible = True
    Else
        Exit For
    End If
Next f

If gNumEntries > -1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False
End If
FlatScrollBar1.Max = gNumEntries
FlatScrollBar1.Value = 0
If gNumEntries <= gMaxLinesOnScreen Then
    FlatScrollBar1.Visible = False
Else
    FlatScrollBar1.Visible = True
End If



End Sub

Private Sub cmdCompleted_DragDrop(Source As Control, X As Single, Y As Single)
Dim f, g
Dim saveSourceData

' If entry is empty (eg after moving last one)
If Trim(lblToDo(Source.Index).Caption) = "" Then
    Exit Sub
End If

saveSourceData = lblToDo(Source.Index).Caption

' Move the ToDos up
For f = Source.Index To gNumEntries - 1
    ' If not last entry on screen
    If f < gMaxLinesOnScreen And f < gNumberShowing Then
        ' move display up 1
        lblToDo(f).Caption = lblToDo(f + 1).Caption
    Else
        ' show next one from array
        Select Case gCurrentFolder
        Case -1   ' main folder
            lblToDo(f).Caption = gMainArray(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gMainArray) - 1
                gMainArray(g) = gMainArray(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gMainArray(UBound(gMainArray) - 1)
            Exit For
        Case 0
            lblToDo(f).Caption = gArray0(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray0) - 1
                gArray0(g) = gArray0(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray0(UBound(gArray0) - 1)
            Exit For
        Case 1
            lblToDo(f).Caption = gArray1(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray2) - 1
                gArray2(g) = gArray2(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray1(UBound(gArray1) - 1)
            Exit For
        Case 2
            lblToDo(f).Caption = gArray2(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray2) - 1
                gArray2(g) = gArray2(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray2(UBound(gArray2) - 1)
            Exit For
        Case 3
            lblToDo(f).Caption = gArray3(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray3) - 1
                gArray3(g) = gArray3(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray3(UBound(gArray3) - 1)
            Exit For
        Case 4
            lblToDo(f).Caption = gArray4(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray4) - 1
                gArray4(g) = gArray4(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray4(UBound(gArray4) - 1)
            Exit For
        Case 5
            lblToDo(f).Caption = gArray5(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray5) - 1
                gArray5(g) = gArray5(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray5(UBound(gArray5) - 1)
            Exit For
        Case 6
            lblToDo(f).Caption = gArray6(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray6) - 1
                gArray6(g) = gArray6(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray6(UBound(gArray6) - 1)
            Exit For
        Case 7
            lblToDo(f).Caption = gArray7(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray7) - 1
                gArray7(g) = gArray7(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray7(UBound(gArray7) - 1)
            Exit For
        Case 8
            lblToDo(f).Caption = gArray8(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray8) - 1
                gArray8(g) = gArray8(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray8(UBound(gArray8) - 1)
            Exit For
        Case 9
            lblToDo(f).Caption = gArray9(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray9) - 1
                gArray9(g) = gArray9(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray9(UBound(gArray9) - 1)
            Exit For
        End Select
    End If
Next f

gNumEntries = gNumEntries - 1
If gNumEntries > 0 Then
    If gNumberShowing < gMaxLinesOnScreen Then
        lblToDo(f).Caption = ""
        lblNumb(f).Caption = ""
        Unload lblNumb(f)
        Unload Line1(f)
        Unload lblToDo(f)
        gNumberShowing = gNumberShowing - 1
        
        If gNumEntries < gMaxLinesOnScreen Then
            FlatScrollBar1.Visible = False
        Else
            FlatScrollBar1.Max = gNumEntries
            FlatScrollBar1.Visible = True
        End If
    End If
Else
    ' clear last entry
    lblToDo(0).Caption = ""
    lblNumb(0).Caption = ""
    gNumberShowing = 0
    FlatScrollBar1.Visible = False
End If

' 3.  ADD SENDER'S DATA TO TARGET (COMPLETED) ARRAY
gCompletedArray(UBound(gCompletedArray)) = saveSourceData
ReDim Preserve gCompletedArray(UBound(gCompletedArray) + 1)

lblActionMsg.Caption = "Transfered to " & cmdCompleted.Caption
cmdSave.Visible = True
End Sub


Private Sub cmdExit_Click()

If frmMain.cmdSave.Visible = True Then
    If MsgBox("Do you want to save the updates first?", 20) = vbYes Then
        Call cmdSave_Click
    End If
End If

Unload Me

End Sub



Private Sub cmdFolder_Click(Index As Integer)
Dim f
On Error Resume Next

' Ignore if same folder
If Index = gCurrentFolder Then
    Exit Sub
End If

gCurrentFolder = Index

' Show delete bin
imgBin.Visible = True
cmdNew.Visible = True
gCurrent1 = 0

' Set FOLDER name
lblCurrentFolder.Caption = cmdFolder(Index).Caption
gCurrentFolder = Index

' CLEAR CURRENT DISPLAY
' For each entry
For f = 1 To gNumberShowing
        ' unLoad lines
        Unload Line1(f)
        ' unLoad Todo boxes
        Unload lblToDo(f)
        ' unLoad Todo numbers
        Unload lblNumb(f)
Next f

' REBUILD DISPLAY ....
gNumEntries = 0
gNumberShowing = 0

' SET ENTRY ONE
lblNumb(0).Caption = "1"
lblToDo(0).Caption = ""
Select Case Index
Case 0
    gNumEntries = UBound(gArray0)
    lblToDo(0).Caption = gArray0(0)
Case 1
    gNumEntries = UBound(gArray1)
    lblToDo(0).Caption = gArray1(0)
Case 2
    gNumEntries = UBound(gArray2)
    lblToDo(0).Caption = gArray2(0)
Case 3
    gNumEntries = UBound(gArray3)
    lblToDo(0).Caption = gArray3(0)
Case 4
    gNumEntries = UBound(gArray4)
    lblToDo(0).Caption = gArray4(0)
Case 5
    gNumEntries = UBound(gArray5)
    lblToDo(0).Caption = gArray5(0)
Case 6
    gNumEntries = UBound(gArray6)
    lblToDo(0).Caption = gArray6(0)
Case 7
    gNumEntries = UBound(gArray7)
    lblToDo(0).Caption = gArray7(0)
Case 8
    gNumEntries = UBound(gArray8)
    lblToDo(0).Caption = gArray8(0)
Case 9
    gNumEntries = UBound(gArray9)
    lblToDo(0).Caption = gArray9(0)
End Select

' SET OTHER ENTRIES
For f = 1 To gNumEntries - 1
    ' If not exceeded screen display
    If f <= gMaxLinesOnScreen Then
        ' Load lines
        Load Line1(f)
        Line1(f).X1 = Line1(0).X1  'left
        Line1(f).X2 = Line1(0).X2  ' right
        Line1(f).Y1 = Line1(f - 1).Y1 + lblToDo(0).Height + 20
        Line1(f).Y2 = Line1(f).Y1
        Line1(f).Visible = True
        ' Load Todo boxes
        Load lblToDo(f)
        lblToDo(f).Top = lblToDo(f - 1).Top + lblToDo(0).Height + 20
        
        Select Case Index
        Case 0
            lblToDo(f).Caption = gArray0(f)
        Case 1
            lblToDo(f).Caption = gArray1(f)
        Case 2
            lblToDo(f).Caption = gArray2(f)
        Case 3
            lblToDo(f).Caption = gArray3(f)
        Case 4
            lblToDo(f).Caption = gArray4(f)
        Case 5
            lblToDo(f).Caption = gArray5(f)
        Case 6
            lblToDo(f).Caption = gArray6(f)
        Case 7
            lblToDo(f).Caption = gArray7(f)
        Case 8
            lblToDo(f).Caption = gArray8(f)
        Case 9
            lblToDo(f).Caption = gArray9(f)
        End Select
        lblToDo(f).Visible = True
        
        ' Load Todo numbers
        Load lblNumb(f)
        lblNumb(f).Top = lblToDo(f).Top
        lblNumb(f).Caption = lblNumb(f - 1).Caption + 1
        lblNumb(f).Visible = True
        
        gNumberShowing = gNumberShowing + 1 ' total showing
    Else
        Exit For
    End If
Next f
If gNumEntries > -1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False
End If
FlatScrollBar1.Max = gNumEntries
If gNumEntries > gMaxLinesOnScreen Then
    FlatScrollBar1.Visible = True
Else
    FlatScrollBar1.Visible = False
End If

End Sub

Private Sub cmdFolder_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Dim f, g
Dim saveSourceData
'
' THIS PROCESSES THE RECEIVING FOLDER
'

' If entry is empty (eg after moving last one)
If Trim(lblToDo(Source.Index).Caption) = "" Then
    Exit Sub
End If

saveSourceData = lblToDo(Source.Index).Caption

' 1. ALTER SENDER'S TODO LIST
' Move the ToDos up
For f = Source.Index To gNumEntries - 1
    ' If screen not full
    If f < gMaxLinesOnScreen And f < gNumberShowing Then
        ' move display up 1
        lblToDo(f).Caption = lblToDo(f + 1).Caption
    Else
        ' 2.  REMOVE FROM SENDER'S ARRAY
        ' fill bottom entry from sender's array
        Select Case gCurrentFolder
        Case -1   ' main folder
            lblToDo(f).Caption = gMainArray(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gMainArray) - 1
                gMainArray(g) = gMainArray(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gMainArray(UBound(gMainArray) - 1)
            Exit For
        Case 0
            lblToDo(f).Caption = gArray0(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray0) - 1
                gArray0(g) = gArray0(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray0(UBound(gArray0) - 1)
            Exit For
        Case 1
            lblToDo(f).Caption = gArray1(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray1) - 1
                gArray1(g) = gArray1(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray1(UBound(gArray1) - 1)
            Exit For
        Case 2
            lblToDo(f).Caption = gArray2(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray2) - 1
                gArray2(g) = gArray2(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray2(UBound(gArray2) - 1)
            Exit For
        Case 3
            lblToDo(f).Caption = gArray3(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray3) - 1
                gArray3(g) = gArray3(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray3(UBound(gArray3) - 1)
            Exit For
        Case 4
            lblToDo(f).Caption = gArray4(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray4) - 1
                gArray4(g) = gArray4(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray4(UBound(gArray4) - 1)
            Exit For
        Case 5
            lblToDo(f).Caption = gArray5(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray5) - 1
                gArray5(g) = gArray5(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray5(UBound(gArray5) - 1)
            Exit For
        Case 6
            lblToDo(f).Caption = gArray6(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray6) - 1
                gArray6(g) = gArray6(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray6(UBound(gArray6) - 1)
            Exit For
        Case 7
            lblToDo(f).Caption = gArray7(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray7) - 1
                gArray7(g) = gArray7(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray7(UBound(gArray7) - 1)
            Exit For
        Case 8
            lblToDo(f).Caption = gArray8(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray8) - 1
                gArray8(g) = gArray8(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray8(UBound(gArray8) - 1)
            Exit For
        Case 9
            lblToDo(f).Caption = gArray9(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray9) - 1
                gArray9(g) = gArray9(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray9(UBound(gArray9) - 1)
            Exit For
        Case 99
            lblToDo(f).Caption = gCompletedArray(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gCompletedArray) - 1
                gCompletedArray(g) = gCompletedArray(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gCompletedArray(UBound(gCompletedArray) - 1)
            Exit For
        End Select
    End If
Next f


gNumEntries = gNumEntries - 1
If gNumEntries > 0 Then
    If gNumberShowing < gMaxLinesOnScreen Then
        lblToDo(f).Caption = ""
        lblNumb(f).Caption = ""
        Unload lblNumb(f)
        Unload Line1(f)
        Unload lblToDo(f)
        gNumberShowing = gNumberShowing - 1
        
        If gNumEntries < gMaxLinesOnScreen Then
            FlatScrollBar1.Visible = False
        Else
            FlatScrollBar1.Max = gNumEntries
            FlatScrollBar1.Visible = True
        End If
    End If
Else
    ' clear last entry
    lblToDo(0).Caption = ""
    lblNumb(0).Caption = ""
    gNumberShowing = 0
    FlatScrollBar1.Visible = False
End If


' Add to the RECEIVING folder's array
f = Index
Select Case f
Case 0
    gArray0(UBound(gArray0)) = saveSourceData
    ReDim Preserve gArray0(UBound(gArray0) + 1)
Case 1
    gArray1(UBound(gArray1)) = saveSourceData
    ReDim Preserve gArray1(UBound(gArray1) + 1)
Case 2
    gArray2(UBound(gArray2)) = saveSourceData
    ReDim Preserve gArray2(UBound(gArray2) + 1)
Case 3
    gArray3(UBound(gArray3)) = saveSourceData
    ReDim Preserve gArray3(UBound(gArray3) + 1)
Case 4
    gArray4(UBound(gArray4)) = saveSourceData
    ReDim Preserve gArray4(UBound(gArray4) + 1)
Case 5
    gArray5(UBound(gArray5)) = saveSourceData
    ReDim Preserve gArray5(UBound(gArray5) + 1)
Case 6
    gArray6(UBound(gArray6)) = saveSourceData
    ReDim Preserve gArray6(UBound(gArray6) + 1)
Case 7
    gArray7(UBound(gArray7)) = saveSourceData
    ReDim Preserve gArray7(UBound(gArray7) + 1)
Case 8
    gArray8(UBound(gArray8)) = saveSourceData
    ReDim Preserve gArray8(UBound(gArray8) + 1)
Case 9
    gArray9(UBound(gArray9)) = saveSourceData
    ReDim Preserve gArray9(UBound(gArray9) + 1)
End Select

lblActionMsg.Caption = "Transfered Task to " & cmdFolder(Index).Caption
cmdSave.Visible = True

End Sub



Private Sub cmdGeneral_Click()
Dim f

' IGNORE IF SAME FOLDER
If lblCurrentFolder.Caption = Right(cmdGeneral.Caption, Len(cmdGeneral.Caption) - 1) Then
    Exit Sub
End If

gCurrentFolder = -1 ' ind General

imgBin.Visible = True
cmdNew.Visible = True
gCurrent1 = 0

' Set FOLDER name
lblCurrentFolder.Caption = Right(cmdGeneral.Caption, Len(cmdGeneral.Caption) - 1)
gCurrentFolder = -1

' unload current display
' For each entry
For f = 1 To gNumberShowing
        ' unLoad lines
        Unload Line1(f)
        ' unLoad Todo boxes
        Unload lblToDo(f)
        ' unLoad Todo numbers
        Unload lblNumb(f)
Next f
gNumberShowing = 0
gNumEntries = UBound(gMainArray)
lblToDo(0).Caption = gMainArray(0)

lblNumb(0).Caption = "1"

' For each entry
For f = 1 To gNumEntries - 1
    ' If not exceeded screen display
    If f <= gMaxLinesOnScreen Then
        ' Load lines
        Load Line1(f)
        Line1(f).X1 = Line1(0).X1  'left
        Line1(f).X2 = Line1(0).X2  ' right
        Line1(f).Y1 = Line1(f - 1).Y1 + lblToDo(0).Height + 20
        Line1(f).Y2 = Line1(f).Y1
        Line1(f).Visible = True
        ' Load Todo boxes
        Load lblToDo(f)
        lblToDo(f).Top = lblToDo(f - 1).Top + lblToDo(0).Height + 20
        lblToDo(f).Caption = gMainArray(f)
        lblToDo(f).Visible = True
        gNumberShowing = gNumberShowing + 1 ' total showing
        ' Load Todo numbers
        Load lblNumb(f)
        lblNumb(f).Top = lblToDo(f).Top
        lblNumb(f).Caption = lblNumb(f - 1).Caption + 1
        
        lblNumb(f).Visible = True
    Else
        Exit For
    End If
Next f

If gNumEntries > -1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False
End If

FlatScrollBar1.Max = gNumEntries
FlatScrollBar1.Value = 0
If gNumEntries <= gMaxLinesOnScreen Then
    FlatScrollBar1.Visible = False
Else
    FlatScrollBar1.Visible = True
End If


End Sub

Private Sub cmdGeneral_DragDrop(Source As Control, X As Single, Y As Single)
Dim f, g
Dim saveSourceData

' PROCESSING WHEN DROPPED ON MAIN FOLDER BUTTON

' MOVE SENDER'S DATA TO MAIN FOLDER (IN ARRAY)
' 1.  ALTER SENDER'S TODO LIST
' 2.  REMOVE FROM SENDER'S ARRAY
' 3.  ADD SENDER'S DATA TO MAIN ARRAY

' If entry is empty (eg after moving last one)
If Trim(lblToDo(Source.Index).Caption) = "" Then
    Exit Sub
End If

' SAVE SENDER'S DATA
saveSourceData = lblToDo(Source.Index).Caption


' 1. ALTER SENDER'S TODO LIST
' Move the ToDos up
For f = Source.Index To gNumEntries - 1
    ' If screen not full
    If f < gMaxLinesOnScreen And f < gNumberShowing Then
        ' move display up 1
        lblToDo(f).Caption = lblToDo(f + 1).Caption
    Else
        ' 2.  REMOVE FROM SENDER'S ARRAY
        ' fill bottom entry from sender's array
        Select Case gCurrentFolder
        Case 0
            lblToDo(f).Caption = gArray0(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray0) - 1
                gArray0(g) = gArray0(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray0(UBound(gArray0) - 1)
            Exit For
        Case 1
            lblToDo(f).Caption = gArray1(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray2) - 1
                gArray2(g) = gArray2(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray1(UBound(gArray1) - 1)
            Exit For
        Case 2
            lblToDo(f).Caption = gArray2(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray2) - 1
                gArray2(g) = gArray2(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray2(UBound(gArray2) - 1)
            Exit For
        Case 3
            lblToDo(f).Caption = gArray3(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray3) - 1
                gArray3(g) = gArray3(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray3(UBound(gArray3) - 1)
            Exit For
        Case 4
            lblToDo(f).Caption = gArray4(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray4) - 1
                gArray4(g) = gArray4(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray4(UBound(gArray4) - 1)
            Exit For
        Case 5
            lblToDo(f).Caption = gArray5(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray5) - 1
                gArray5(g) = gArray5(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray5(UBound(gArray5) - 1)
            Exit For
        Case 6
            lblToDo(f).Caption = gArray6(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray6) - 1
                gArray6(g) = gArray6(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray6(UBound(gArray6) - 1)
            Exit For
        Case 7
            lblToDo(f).Caption = gArray7(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray7) - 1
                gArray7(g) = gArray7(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray7(UBound(gArray7) - 1)
            Exit For
        Case 8
            lblToDo(f).Caption = gArray8(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray8) - 1
                gArray8(g) = gArray8(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray8(UBound(gArray8) - 1)
            Exit For
        Case 9
            lblToDo(f).Caption = gArray9(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gArray9) - 1
                gArray9(g) = gArray9(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gArray9(UBound(gArray9) - 1)
            Exit For
        Case 99
            lblToDo(f).Caption = gCompletedArray(f + gCurrent1 + 1)
            ' alter array
            For g = Source.Index + gCurrent1 To UBound(gCompletedArray) - 1
                gCompletedArray(g) = gCompletedArray(g + 1)
            Next g
            ' remove last entry
            ReDim Preserve gCompletedArray(UBound(gCompletedArray) - 1)
            Exit For
    
        End Select
    End If
Next f

gNumEntries = gNumEntries - 1
If gNumEntries > 0 Then
    If gNumberShowing < gMaxLinesOnScreen Then
        lblToDo(f).Caption = ""
        lblNumb(f).Caption = ""
        Unload lblNumb(f)
        Unload Line1(f)
        Unload lblToDo(f)
        gNumberShowing = gNumberShowing - 1
        
        If gNumEntries < gMaxLinesOnScreen Then
            FlatScrollBar1.Visible = False
        Else
            FlatScrollBar1.Max = gNumEntries
            FlatScrollBar1.Visible = True
        End If
    End If
Else
    ' clear last entry
    lblToDo(0).Caption = ""
    lblNumb(0).Caption = ""
    gNumberShowing = 0
    FlatScrollBar1.Visible = False
End If


' 3.  ADD SENDER'S DATA TO TARGET (MAIN) ARRAY
gMainArray(UBound(gMainArray)) = saveSourceData
ReDim Preserve gMainArray(UBound(gMainArray) + 1)

lblActionMsg.Caption = "Transfered to " & cmdGeneral.Caption
cmdSave.Visible = True
End Sub


Private Sub cmdMaintain_Click()

' Save current
Call cmdSave_Click

Form2.Show 1

Call locFolders
Call locFillArrays

End Sub

Private Sub cmdNew_Click()
 
cmdAddEnd.Visible = True
cmdAddStart.Visible = True
cmdNew.Visible = False
cmdCancel.Visible = True
cmdSave.Visible = True

txtEntry.Visible = True
txtEntry.SetFocus

End Sub

Private Sub cmdPrint_Click()

Dim f

frmPreview.List1.AddItem ""
frmPreview.List1.AddItem "TO DO Listing from folder: " & lblCurrentFolder.Caption
frmPreview.List1.AddItem ""

frmPreview.List1.AddItem "As at " & Format(Now, "dddd dd mmm yyyy, hh:mmam/pm")
frmPreview.List1.AddItem ""

frmPreview.List1.AddItem "No.  Item"
frmPreview.List1.AddItem ""

For f = 0 To gNumEntries - 1
    frmPreview.List1.AddItem Format(f + 1, "00") & "   " & lblToDo(f)
Next f

frmPreview.Show 1

End Sub

Private Sub cmdSave_Click()

' output temp new file
' delete previous
' rename new to .dat
On Error Resume Next
Dim f

Screen.MousePointer = vbHourglass
gFile2 = FreeFile

Open gAppPath & "todos.tmp" For Output As gFile2
If Err = 0 Then
    ' Store entries from array(s)
    For f = 0 To UBound(gMainArray) - 1
        If Trim(gMainArray(f)) <> "" Then
            Print #gFile2, "Folder1" & "," & gMainArray(f)
        End If
    Next f
    For f = 0 To UBound(gArray0) - 1
        If Trim(gArray0(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 0) & "," & gArray0(f)
        End If
    Next f
    For f = 0 To UBound(gArray1) - 1
        If Trim(gArray1(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 1) & "," & gArray1(f)
        End If
    Next f
    For f = 0 To UBound(gArray2) - 1
        If Trim(gArray2(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 2) & "," & gArray2(f)
        End If
    Next f
    For f = 0 To UBound(gArray3) - 1
        If Trim(gArray3(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 3) & "," & gArray3(f)
        End If
    Next f
    For f = 0 To UBound(gArray4) - 1
        If Trim(gArray4(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 4) & "," & gArray4(f)
        End If
    Next f
    For f = 0 To UBound(gArray5) - 1
        If Trim(gArray5(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 5) & "," & gArray5(f)
        End If
    Next f
    For f = 0 To UBound(gArray6) - 1
        If Trim(gArray6(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 6) & "," & gArray6(f)
        End If
    Next f
    For f = 0 To UBound(gArray7) - 1
        If Trim(gArray7(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 7) & "," & gArray7(f)
        End If
    Next f
    For f = 0 To UBound(gArray8) - 1
        If Trim(gArray8(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 8) & "," & gArray8(f)
        End If
    Next f
    For f = 0 To UBound(gArray9) - 1
        If Trim(gArray9(f)) <> "" Then
            Print #gFile2, gSubArrays(0, 9) & "," & gArray9(f)
        End If
    Next f
    For f = 0 To UBound(gCompletedArray) - 1
        If Trim(gCompletedArray(f)) <> "" Then
            Print #gFile2, "COMP" & "," & gCompletedArray(f)
        End If
    Next f
    
End If

Close gFile1

FileCopy gAppPath & "todos.dat", gAppPath & "\backups\todos" & Format(Now, "yymmddhhmm") & ".dat"
Kill gAppPath & "todos.dat"

Close gFile2
Name gAppPath & "todos.tmp" As gAppPath & "todos.dat"

Screen.MousePointer = vbDefault
cmdSave.Visible = False
End Sub

Private Sub FlatScrollBar1_Change()

Dim f

''MsgBox FlatScrollBar1.Value
lblToDo(0).Caption = ""
' Alter display ..........
' For each entry
For f = 0 To gNumEntries - 1

    ' If not exceeded screen display and some left in Array
    If f <= gMaxLinesOnScreen And (f + FlatScrollBar1.Value < gNumEntries + 1) Then
        
        ' If slot not loaded
        If f <> 0 And f > gNumberShowing Then
            ' Load a line
            Load Line1(f)
            Line1(f).X1 = Line1(0).X1  'left
            Line1(f).X2 = Line1(0).X2  ' right
            Line1(f).Y1 = Line1(f - 1).Y1 + lblToDo(0).Height + 20
            Line1(f).Y2 = Line1(f).Y1
            Line1(f).Visible = True
            ' Load Todo boxes
            Load lblToDo(f)
            lblToDo(f).Top = lblToDo(f - 1).Top + lblToDo(0).Height + 20
            lblToDo(f).Visible = True
            gNumberShowing = gNumberShowing + 1 ' total showing
            ' Load Todo numbers
            Load lblNumb(f)
            lblNumb(f).Top = lblToDo(f).Top
            lblNumb(f).Caption = lblNumb(f - 1).Caption + 1
            lblNumb(f).Visible = True
        End If
        
        ' Show entry
        Select Case gCurrentFolder
        Case -1
            lblToDo(f).Caption = gMainArray(f + FlatScrollBar1.Value)
        Case 0
            lblToDo(f).Caption = gArray0(f + FlatScrollBar1.Value)
        Case 1
            lblToDo(f).Caption = gArray1(f + FlatScrollBar1.Value)
        Case 2
            lblToDo(f).Caption = gArray2(f + FlatScrollBar1.Value)
        Case 3
            lblToDo(f).Caption = gArray3(f + FlatScrollBar1.Value)
        Case 4
            lblToDo(f).Caption = gArray4(f + FlatScrollBar1.Value)
        Case 5
            lblToDo(f).Caption = gArray5(f + FlatScrollBar1.Value)
        Case 6
            lblToDo(f).Caption = gArray6(f + FlatScrollBar1.Value)
        Case 7
            lblToDo(f).Caption = gArray7(f + FlatScrollBar1.Value)
        Case 8
            lblToDo(f).Caption = gArray8(f + FlatScrollBar1.Value)
        Case 9
            lblToDo(f).Caption = gArray9(f + FlatScrollBar1.Value)

        End Select
        
        ' Update Todo numbers
        lblNumb(f).Caption = FlatScrollBar1.Value + f + 1
        
    Else
        ' Display full or all showing
        If FlatScrollBar1.Value > gCurrent1 Then
            ' scrolling up - see if at end
            Do While f <= gMaxLinesOnScreen
                If f <= gNumberShowing Then
                    lblToDo(f).Caption = ""
                    lblNumb(f).Caption = ""
                    Unload Line1(f)
                    Unload lblToDo(f)
                    Unload lblNumb(f)
                End If
                f = f + 1
            Loop
        End If
        gNumberShowing = gNumEntries - FlatScrollBar1.Value
        If gNumberShowing > gMaxLinesOnScreen Then
            gNumberShowing = gMaxLinesOnScreen
        End If
        gCurrent1 = FlatScrollBar1.Value
        
        Exit Sub  '====================>>>>>
        
    End If
Next f

'If f > gNumEntries + 1 And f > gNumberShowing And gNumberShowing <> gMaxLinesOnScreen Then
'If gNumberShowing <> gMaxLinesOnScreen Then

If lblToDo.Count > 1 And lblToDo.Count > f Then
' remove a line from display
    lblToDo(f).Caption = ""
    lblNumb(f).Caption = ""
    Unload Line1(f)
    Unload lblToDo(f)
    Unload lblNumb(f)
End If

gCurrent1 = FlatScrollBar1.Value
 
End Sub

Private Sub FlatScrollBar1_Scroll()
Call FlatScrollBar1_Change
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 34 Then ' page down
    If FlatScrollBar1.Value > FlatScrollBar1.Max - gMaxLinesOnScreen Then
        FlatScrollBar1.Value = FlatScrollBar1.Max
    Else
        FlatScrollBar1.Value = FlatScrollBar1.Value + gMaxLinesOnScreen
    End If
End If

If KeyCode = 33 Then ' page up
    If FlatScrollBar1.Value < gMaxLinesOnScreen Then
        FlatScrollBar1.Value = 0
    Else
        FlatScrollBar1.Value = FlatScrollBar1.Value - gMaxLinesOnScreen
    End If
End If

End Sub

Private Sub Form_Load()
Dim f, FileNum, fName, fNumb
Dim both, folder, details, matched
Dim fileRec

On Error Resume Next    ' in case problem with .dat files
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

' Set path to .dat files
gAppPath = App.Path
If Right(gAppPath, 1) <> "\" Then
    gAppPath = gAppPath & "\"
End If
 
' Set FOLDER name
lblCurrentFolder.Caption = Right(cmdGeneral.Caption, Len(cmdGeneral.Caption) - 1)

' Limit the number of entries on screen
gMaxLinesOnScreen = 17  ' Base 0

' Initialise number of 1st entry on screen (Base 0)
gCurrent1 = 0

' Inititialise  array of entries
ReDim gMainArray(0)
ReDim gSubArrays(1, 0) ' 0=name 1=array number
ReDim gCompletedArray(0)

Call locFolders

gCurrentFolder = -1  ' Main general

Call locFillArrays

If gNumEntries > -1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False
End If
FlatScrollBar1.Max = gNumEntries
FlatScrollBar1.Value = 0
If gNumEntries <= gMaxLinesOnScreen Then
    FlatScrollBar1.Visible = False
Else
    FlatScrollBar1.Visible = True
End If

' Position 1st separator line
Line1(0).X1 = 60    ' left
Line1(0).X2 = lblToDo(0).Width + 660 ' right
Line1(0).Y1 = lblToDo(0).Top + lblToDo(0).Height + 5 ' top
Line1(0).Y2 = Line1(0).Y1  ' bottom!

' For each entry
For f = 1 To gNumEntries - 1
    ' If not exceeded screen display
    If f <= gMaxLinesOnScreen Then
        ' Load lines
        Load Line1(f)
        Line1(f).X1 = Line1(0).X1  'left
        Line1(f).X2 = Line1(0).X2  ' right
        Line1(f).Y1 = Line1(f - 1).Y1 + lblToDo(0).Height + 20
        Line1(f).Y2 = Line1(f).Y1
        Line1(f).Visible = True
        ' Load Todo boxes
        Load lblToDo(f)
        lblToDo(f).Top = lblToDo(f - 1).Top + lblToDo(0).Height + 20
        lblToDo(f).Caption = gMainArray(f)
        lblToDo(f).Visible = True
        gNumberShowing = gNumberShowing + 1 ' total showing
        ' Load Todo numbers
        Load lblNumb(f)
        lblNumb(f).Top = lblToDo(f).Top - 20
        lblNumb(f).Caption = lblNumb(f - 1).Caption + 1
        lblNumb(f).Visible = True
    Else
        Exit For
    End If
Next f


End Sub



Private Sub imgBin_DragDrop(Source As Control, X As Single, Y As Single)

Dim f
If MsgBox("OK to delete entry" & Chr(10) & Chr(10) & Source.Caption & "?", 36) = vbYes Then
    
    ' Rebuild array, per Folder Number
    Select Case gCurrentFolder
    Case -1
         ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gMainArray) - 1
            gMainArray(f) = gMainArray(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gMainArray(UBound(gMainArray) - 1)
    Case 0
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray0) - 1
            gArray0(f) = gArray0(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray0(UBound(gArray0) - 1)
    Case 1
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray1) - 1
            gArray1(f) = gArray1(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray1(UBound(gArray1) - 1)
    Case 2
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray2) - 1
            gArray2(f) = gArray2(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray2(UBound(gArray2) - 1)
    Case 3
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray3) - 1
            gArray3(f) = gArray3(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray3(UBound(gArray3) - 1)
    Case 4
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray4) - 1
            gArray4(f) = gArray4(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray4(UBound(gArray4) - 1)
    Case 5
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray5) - 1
            gArray5(f) = gArray5(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray5(UBound(gArray5) - 1)
    Case 6
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray6) - 1
            gArray6(f) = gArray6(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray6(UBound(gArray6) - 1)
    Case 7
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray7) - 1
            gArray7(f) = gArray7(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray7(UBound(gArray7) - 1)
    Case 8
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray8) - 1
            gArray8(f) = gArray8(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray8(UBound(gArray8) - 1)
    Case 9
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gArray9) - 1
            gArray9(f) = gArray9(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gArray9(UBound(gArray9) - 1)
    Case 99
        ' Move array entries up one
        For f = Source.Index + gCurrent1 To UBound(gCompletedArray) - 1
            gCompletedArray(f) = gCompletedArray(f + 1)
        Next f
        ' Decrease array size
        ReDim Preserve gCompletedArray(UBound(gCompletedArray) - 1)
    End Select
    
    gNumEntries = gNumEntries - 1
    ' Display
    If gNumEntries < gMaxLinesOnScreen Then
        gNumberShowing = gNumberShowing - 1  ' BASE 0
    End If
    If FlatScrollBar1.Value = 0 Then
        Call FlatScrollBar1_Change
    Else
        FlatScrollBar1.Value = 0
    End If
    
    cmdSave.Visible = True
End If

End Sub


Private Sub lblToDo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Dim f
Dim saveSourceData
Dim saveTargetData

' No movement
If Source.Index = Index Then
    Exit Sub
End If

' If preceded by a right-click edit
If txtEntry.Visible = True Then
    If MsgBox("Ignore previous edit?", 36) = vbNo Then
        txtEntry.SetFocus
        Exit Sub
    Else
        Call cmdCancel_Click
    End If
End If

saveSourceData = lblToDo(Source.Index).Caption
saveTargetData = lblToDo(Index).Caption

If Source.Index < Index Then
    ' Move the ToDos up
    For f = Source.Index To Index - 1
        lblToDo(f).Caption = lblToDo(f + 1).Caption
        
        Select Case gCurrentFolder
            Case -1
                gMainArray(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 0
                gArray0(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 1
                gArray1(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 2
                gArray2(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 3
                gArray3(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 4
                gArray4(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 5
                gArray5(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 6
                gArray6(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 7
                gArray7(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 8
                gArray8(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 9
                gArray9(f + gCurrent1) = lblToDo(f + 1).Caption
            Case 99
                gCompletedArray(f + gCurrent1) = lblToDo(f + 1).Caption
        End Select
        
    Next f
Else
    ' Move the ToDos down
    For f = Source.Index To Index + 1 Step -1
        lblToDo(f).Caption = lblToDo(f - 1).Caption
        
        Select Case gCurrentFolder
            Case -1
                gMainArray(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 0
                gArray0(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 1
                gArray1(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 2
                gArray2(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 3
                gArray3(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 4
                gArray4(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 5
                gArray5(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 6
                gArray6(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 7
                gArray7(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 8
                gArray8(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 9
                gArray9(f + gCurrent1) = lblToDo(f - 1).Caption
            Case 99
                gCompletedArray(f + gCurrent1) = lblToDo(f - 1).Caption

        End Select
    Next f
End If

' Move the box to new position
''''''lblToDo(Source.Index).Top = lblToDo(Index).Top
lblToDo(Index).Caption = saveSourceData

Select Case gCurrentFolder
    Case -1
        gMainArray(f + gCurrent1) = saveSourceData
    Case 0
        gArray0(f + gCurrent1) = saveSourceData
    Case 1
        gArray1(f + gCurrent1) = saveSourceData
    Case 2
        gArray2(f + gCurrent1) = saveSourceData
    Case 3
        gArray3(f + gCurrent1) = saveSourceData
    Case 4
        gArray4(f + gCurrent1) = saveSourceData
    Case 5
        gArray5(f + gCurrent1) = saveSourceData
    Case 6
        gArray6(f + gCurrent1) = saveSourceData
    Case 7
        gArray7(f + gCurrent1) = saveSourceData
    Case 8
        gArray8(f + gCurrent1) = saveSourceData
    Case 9
        gArray9(f + gCurrent1) = saveSourceData
    Case 99
        gCompletedArray(f + gCurrent1) = saveSourceData
End Select

lblActionMsg.Caption = "Moved Task from " & gCurrent1 + Source.Index + 1 & " to " & gCurrent1 + Index + 1

cmdSave.Visible = True
End Sub


Private Sub lblToDo_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
Source.DragIcon = imgMoveIt.Picture

End Sub


Private Sub lblToDo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a
If cmdNew.Visible = False Then
    Exit Sub
End If

' Right click to amend
If Button = 2 Then
    ' save index
    gSelIndex = Index
    txtEntry.Text = lblToDo(Index).Caption
    txtEntry.Visible = True
    cmdAmend.Visible = True
    cmdCancel.Visible = True
    cmdNew.Visible = False
    txtEntry.SelStart = Len(txtEntry.Text)
    txtEntry.SetFocus
End If


End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)

' If ENTER key
If KeyAscii = 13 Then
    If cmdAmend.Visible Then
        ' AMENDING
        Call cmdAmend_Click
    Else
        ' ADDING
        Call cmdAddEnd_Click
    End If
End If

End Sub



Public Sub locFolders()
Dim f, FileNum, fName, fNumb

ReDim gSubArrays(1, 0)

If cmdFolder.Count <> 0 Then
    For f = 1 To cmdFolder.Count - 1
        Unload cmdFolder(f)
    Next f
End If

FileNum = FreeFile
f = 0

Open gAppPath & "folders.dat" For Input As FileNum
If Err = 0 Then
    Do While Not EOF(FileNum)
        Input #FileNum, fName
        gSubArrays(0, f) = fName ' name
        gSubArrays(1, f) = f     ' array number
        ' Show buttons
        If f = 0 Then
           cmdFolder(0).Caption = fName
           cmdFolder(0).Visible = True
        Else
           Load cmdFolder(f)
           cmdFolder(f).Left = cmdFolder(0).Left
           cmdFolder(f).Top = cmdFolder(f - 1).Top + cmdFolder(0).Height
           cmdFolder(f).Caption = fName
           cmdFolder(f).Visible = True
           ' Set TAB INDEX
           cmdFolder(f).TabIndex = cmdFolder(f - 1).TabIndex + 1
        End If
        ReDim Preserve gSubArrays(1, UBound(gSubArrays, 2) + 1)
        f = f + 1
    Loop
End If
Err = 0
Close FileNum

End Sub

Public Sub locFillArrays()
Dim f, FileNum, fName, fNumb, both, folder, details, matched

ReDim gMainArray(0)
ReDim gCompletedArray(0)
ReDim gArray0(0)
ReDim gArray1(0)
ReDim gArray2(0)
ReDim gArray3(0)
ReDim gArray4(0)
ReDim gArray5(0)
ReDim gArray6(0)
ReDim gArray7(0)
ReDim gArray8(0)
ReDim gArray9(0)
 
gNumEntries = 0
gFile1 = FreeFile
Err = 0 ' in case Folder file not found previously
Open gAppPath & "todos.dat" For Input As gFile1
If Err = 0 Then
    ' Store entries in array
    Do While Not EOF(gFile1)
        Line Input #gFile1, both ' folder, details
        folder = Left(both, InStr(both, ",") - 1)
        details = Right(both, Len(both) - InStr(both, ","))
        ' If main folder
        If UCase(folder) = "FOLDER1" Then
            gMainArray(UBound(gMainArray)) = details
            ReDim Preserve gMainArray(UBound(gMainArray) + 1)
        ElseIf UCase(folder) = "COMP" Then
            gCompletedArray(UBound(gCompletedArray)) = details
            ReDim Preserve gCompletedArray(UBound(gCompletedArray) + 1)
        Else
            ' Build array of other folder names
            ' See if new name
            matched = False
            For f = 0 To UBound(gSubArrays, 2)
                If UCase(gSubArrays(0, f)) = UCase(folder) Then
                    matched = True
                    Exit For ' leave f set
                End If
            Next f
            
            If Not matched Then  ' f-1 = next
                f = f - 1
                gSubArrays(0, f) = folder  ' name
                gSubArrays(1, f) = f     ' array number
                ReDim Preserve gSubArrays(1, UBound(gSubArrays, 2) + 1)
                
                ' Show buttons
                If f = 0 Then
                   cmdFolder(0).Caption = folder
                   cmdFolder(0).Visible = True
                Else
                   Load cmdFolder(f)
                   cmdFolder(f).Left = cmdFolder(0).Left
                   cmdFolder(f).Top = cmdFolder(f - 1).Top + cmdFolder(0).Height
                   cmdFolder(f).Caption = folder
                   cmdFolder(f).Visible = True
                End If
 
                ' UPDATE THE FOLDER.DAT FILE
                FileNum = FreeFile
                Open gAppPath & "folders.dat" For Output As FileNum
                For f = 0 To UBound(gSubArrays, 2)
                    Print #FileNum, gSubArrays(0, f)
                Next f
                Close FileNum
            End If
         
            ' Store entry in the correct sub array (f)
            Select Case f
            Case 0
                gArray0(UBound(gArray0)) = details
                ReDim Preserve gArray0(UBound(gArray0) + 1)
            Case 1
                gArray1(UBound(gArray1)) = details
                ReDim Preserve gArray1(UBound(gArray1) + 1)
            Case 2
                gArray2(UBound(gArray2)) = details
                ReDim Preserve gArray2(UBound(gArray2) + 1)
            Case 3
                gArray3(UBound(gArray3)) = details
                ReDim Preserve gArray3(UBound(gArray3) + 1)
            Case 4
                gArray4(UBound(gArray4)) = details
                ReDim Preserve gArray4(UBound(gArray4) + 1)
            Case 5
                gArray5(UBound(gArray5)) = details
                ReDim Preserve gArray5(UBound(gArray5) + 1)
            Case 6
                gArray6(UBound(gArray6)) = details
                ReDim Preserve gArray6(UBound(gArray6) + 1)
            Case 7
                gArray7(UBound(gArray7)) = details
                ReDim Preserve gArray7(UBound(gArray7) + 1)
            Case 8
                gArray8(UBound(gArray8)) = details
                ReDim Preserve gArray8(UBound(gArray8) + 1)
            Case 9
                gArray9(UBound(gArray9)) = details
                ReDim Preserve gArray9(UBound(gArray9) + 1)
            End Select
        
        End If
    Loop
    lblToDo(0).Caption = gMainArray(0)
    gNumEntries = UBound(gMainArray)
End If

End Sub
