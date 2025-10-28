VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RelocateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'обработка нажатия Отмена
Private Sub cancelButton_Click()

   DoCmd.Close acForm, "RelocateForm"
   cancelUpdateQwr = True
   
End Sub 'cancelButton_Click
'процедура выполняемая при загрузке формы
Private Sub Form_Load()

   recordsNumberTextBox.SetFocus
   recordsNumberTextBox.Text = rl.recsCount
   recordsNumberTextBox.Locked = True
   relocationDateTextBox.SetFocus
   
End Sub 'Form_Load

'обработка нажатия ОК
Private Sub okButton_Click()
Dim rD As String, pers As String, number As String, reason As String, state As String
On Error GoTo ErrHandler

   'вводим, проверяем и обновляем информацию в соответствующих полях формы UpdateСклад
   'после чего снова их блокируем
   
   relocationDateTextBox.SetFocus
   rD = relocationDateTextBox.Text
   If rD <> "" Then
      If CDate(rD) <= Date Then
         Forms("UpdateСклад").relocationDateTextBox.SetFocus
         Forms("UpdateСклад").relocationDateTextBox.Locked = False
'         Forms("UpdateСклад").relocationDateTextBox.Text = rD
         Forms("UpdateСклад").relocationDateTextBox.Locked = True
      Else
         MsgBox "Введите дату не позднее текущей!", vbExclamation, strMsgBox
         Exit Sub
      End If
   Else
      MsgBox "Отсутствует дата", vbExclamation, strMsgBox
      Exit Sub
   End If
   
   personalTextBox.SetFocus
   pers = personalTextBox.Text
   If pers <> "" Then
      Forms("UpdateСклад").personalTextBox.SetFocus
      Forms("UpdateСклад").personalTextBox.Locked = False
      Forms("UpdateСклад").personalTextBox.Text = pers
      Forms("UpdateСклад").personalTextBox.Locked = True
   Else
      MsgBox "Отсутствует фамилия", vbExclamation, strMsgBox
      Exit Sub
   End If
   
   reasonComboBox.SetFocus
   reason = reasonComboBox.Text
   If reason <> "" Then
      Forms("UpdateСклад").reasonComboBox.SetFocus
      Forms("UpdateСклад").reasonComboBox.Locked = False
      'Forms("UpdateСклад").reasonComboBox.Text = reason
      Forms("UpdateСклад").reasonComboBox.Locked = True
   Else
      MsgBox "Отсутствует основание перемещения", vbExclamation, strMsgBox
      Exit Sub
   End If
   
   stateComboBox.SetFocus
   state = stateComboBox.Text
   If state <> "" Then
      Forms("UpdateСклад").stateComboBox.SetFocus
      Forms("UpdateСклад").stateComboBox.Locked = False
      'Forms("UpdateСклад").stateComboBox.Text = state
      Forms("UpdateСклад").stateComboBox.Locked = True
   Else
      MsgBox "Отсутствует состояние при перемещении", vbExclamation, strMsgBox
      Exit Sub
   End If
   
   numberTextBox.SetFocus
   number = numberTextBox.Text
   If number <> "" Then
   '*******************************************
   '     заменяемый блок
      recordsNumberTextBox.SetFocus
      If Val(number) <= Val(recordsNumberTextBox.Text) Then
         Forms("UpdateСклад").relocNumberTextBox.SetFocus
         Forms("UpdateСклад").relocNumberTextBox.Locked = False
         Forms("UpdateСклад").relocNumberTextBox.Text = number
         Forms("UpdateСклад").relocNumberTextBox.Locked = True
   
      '   If Val(Forms("UpdateСклад").Количество.Text) = 0 Then
      '      Forms("UpdateСклад").Количество.Text = CStr(Val(Forms("UpdateСклад").Количество.Text) +
      '_
      '         Val(number))
      '   ElseIf Val(Forms("UpdateСклад").Количество.Text) > 0 Then
      '      Forms("UpdateСклад").Количество.Text = CStr(Val(Forms("UpdateСклад").Количество.Text) -
      '_
      '         Val(number))
      '   End If
      '   Forms("UpdateСклад").Количество.Locked = True
      
      '     заменяемый блок
      '*******************************************
      Else
         MsgBox "Можно переместить не более " & recordsNumberTextBox.Text, vbExclamation, strMsgBox
         Exit Sub
      End If
   Else
      MsgBox "Отсутствует количество", vbExclamation, strMsgBox
      Exit Sub
   End If
   
   DoCmd.Close acForm, "RelocateForm"
   
Exit Sub
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description & " в процедуре Private Sub relocateButton_Click формы RelocateForm"
   Resume Next
End Sub 'okButton_Click
