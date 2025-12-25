Attribute VB_Name = "modExport"
Option Explicit

Public Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String

    exportPath = ThisWorkbook.Path & "\src\"

    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
        End Select
    Next vbComp

    MsgBox "Export complete!", vbInformation
End Sub

