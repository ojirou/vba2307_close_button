Attribute VB_Name = "Module1"
Private Sub CommandButton2_Click_1()
    ret = MsgBox("�I�����܂��A��낵���ł����H", vbYesNo, "�m�F")
    If ret = 6 Then
        ActiveWorkbook.Close SaveChanges:=False
        DoEvents
    End If
End Sub

