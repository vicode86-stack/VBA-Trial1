VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    Dim a As Double
    Dim b As Double
    Dim result As Double

    ' Validate input
    If Not IsNumeric(txtA.value) Or Not IsNumeric(txtB.value) Then
        MsgBox "Please enter valid numbers", vbExclamation
        Exit Sub
    End If

    a = CDbl(txtA.value)
    b = CDbl(txtB.value)

    ' Determine formula
    If optAddSq.value = True Then
        result = (a + b) ^ 2

    ElseIf optSubSq.value = True Then
        result = (a - b) ^ 2

    ElseIf optSumSq.value = True Then
        result = a ^ 2 + b ^ 2

    Else
        MsgBox "Please select a formula", vbExclamation
        Exit Sub
    End If

    lblResult.Caption = "Result: " & result
End Sub

