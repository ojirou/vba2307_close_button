Attribute VB_Name = "Module1"
Private Sub CommandButton2_Click_1()
    ret = MsgBox("終了します、よろしいですか？", vbYesNo, "確認")
    If ret = 6 Then
        ActiveWorkbook.Close SaveChanges:=False
        DoEvents
    End If
End Sub

