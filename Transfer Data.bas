Attribute VB_Name = "Module3"
Sub Transfer()
Attribute Transfer.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' Transfer Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    Sheets("Table").Select
    Range("Invoices[[#Headers],[Invoice Number]]").Select
    Selection.End(xlDown).Select
    Selection.ListObject.ListRows.Add AlwaysInsert:=True
    ActiveCell.Offset(1, 0).Range("Invoices[[#Headers],[Invoice Number]]").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+1"
    ActiveCell.Select
    Sheets("Data Entry").Select
    Range("C5").Select
    Selection.Copy
    Sheets("Table").Select
    ActiveCell.Offset(0, 1).Range("Invoices[[#Headers],[Invoice Number]]").Select
    ActiveSheet.Paste
    Sheets("Data Entry").Select
    Range("C7:C9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Table").Select
    ActiveCell.Offset(0, 1).Range("Invoices[[#Headers],[Invoice Number]]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    ActiveCell.Offset(0, 6).Range("Invoices[[#Headers],[Invoice Number]]").Select
    Sheets("Data Entry").Select
    Range("C7:C9,C5").Select
    Range("C5").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("C5").Select
End Sub
