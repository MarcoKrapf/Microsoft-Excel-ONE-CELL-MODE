Attribute VB_Name = "Modul1"
Option Explicit

Public Sub OneCellMode(control As IRibbonControl)
    If MsgBox("Do you really want to activate the MICROSOFT EXCEL ONE CELL MODE (also called FLEUERER MODE or ABSOLUTE BEGINNER MODE) ?!", vbYesNo) = vbYes Then
        With Range("A1:XFD1048576")
            .Clear
            .Merge
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        MsgBox "You can now start editing cell A1."
    End If
End Sub
