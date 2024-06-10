Attribute VB_Name = "Module5"
Sub Issue_Invoice()
Attribute Issue_Invoice.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' Issue_Invoice Macro
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Sheets("Invoice Template").Select
    Range("B3").Select
    ActiveCell.FormulaR1C1 = InputBox("Please enter the invoice number", "Issue Invoice")
    Range("B4").Select
End Sub
