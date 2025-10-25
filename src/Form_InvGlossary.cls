VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InvGlossary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' кн. Выход
Private Sub closeButton_Click()

   DoCmd.Close acForm, "InvGlossary"

End Sub ' closeButton_Click()

' перед обновлением данных
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrHandler
Dim strMsg As String
Dim iResponse As Integer

   ' Текст выводимого сообщения.
   strMsg = "Вы желаете сохранить изменения ?" & Chr(10)
   strMsg = strMsg & "Нажмите Да для сохранения или Нет для отмены сделанных изменений."

   ' Вывод сообщения.
   iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Сохранить запись?")
   
   ' Проверить ответ пользователя.
   If iResponse = vbNo Then
   
      ' Откатить изменения.
      DoCmd.RunCommand acCmdUndo

      ' Отменить обновление записи.
      Cancel = True
   End If

Exit Sub
ErrHandler:
   MsgBox "Error #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Me.Undo
   Resume Next
   
End Sub ' Form_BeforeUpdate
' кн. Сохранить
Private Sub saveButton_Click()

   Me.Refresh
   
End Sub ' saveButton_Click()
