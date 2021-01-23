Attribute VB_Name = "Module1"
Option Explicit

Global gAppPath         ' path to data (.dat) files

Global gNumEntries      ' 1 + num of active entries

Global gMaxLinesOnScreen  ' max entries on page
                        ' for scroll bar visiblity
Global gCurrentFolder
                        
Global gMainArray()     ' main array of entries
Global gSubArrays()     ' 0=name 1=array number
Global gArray0()
Global gArray1()
Global gArray2()
Global gArray3()
Global gArray4()
Global gArray5()
Global gArray6()
Global gArray7()
Global gArray8()
Global gArray9()
Global gCompletedArray()

Global gCurrent1        ' number of 1st entry on screen
Global gNumberShowing   ' number of entries on screen

Global gFile1
Global gFile2

Global gSelIndex        ' Amend right click

Global gCancelPrinting

