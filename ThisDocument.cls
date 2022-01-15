'Copyright Moskalev Dmitry
Sub CommandButton1_Click()
ActiveDocument.UpdateStylesOnOpen = False
Application.Visible = False

If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
End If

If (ActiveWindow.View.Type <> wdPrintView) Then
    ActiveWindow.View.Type = wdPrintView
End If

If (AddIns.Count >= 1) Then
MsgBox "Please turn off the addons:" & AddIns.Count
Application.Quit SaveChanges:=wdDoNotSaveChanges
End If

If (Word.Documents.Count > 1) Then
MsgBox "Please close other MS Word documents"
ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
End If

If (ActiveDocument.ProtectionType <> wdAllowOnlyReading) Then
Application.Quit SaveChanges:=wdDoNotSaveChanges
End If

If (Application.Version >= "12") Then
UserForm1.Show
Else
ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
End If
End Sub
