VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Authorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'объявление источника данных для запроса филиала, логина и пароля из таблицы "Структура"
Dim rst As Recordset

'обработка нажатия Отмена
Private Sub cancelButton_Click()

   Application.Quit acQuitSaveNone
   
End Sub 'cancelButton_Click

'запуск в режиме конструирования
Private Sub debugModeButton_Click()

   auth.setSubdiv ("Снежное")
   auth.setLogin ("admin8")
   auth.setPassword ("admin8")
   'экземпляр класса Recordset
            Set rst = CurrentDb.OpenRecordset("SELECT Код " _
               & "FROM Структура " _
               & " WHERE Перевод LIKE '" _
               & "*" & auth.getSubdiv & "*'" & "")
   rst.MoveFirst
   auth.setId (rst.Fields(0))
   Forms("Authorization").Visible = False
   DoCmd.OpenForm "ViewСклад", acNormal, , , acFormAdd, acWindowNormal

   'clean
   rst.Close
   
End Sub 'debugModeButton_Click

Private Sub Form_Load()

   'конструктор нового экземпляра класса clsAuthorizationClass хранения и выдачи по запросу регистрационных данных
   'текущего пользователя
   Set auth = New clsAuthorizationClass
   Call initVarsInMain
   
End Sub
'создание внешних ключей в столбце Тип таблицы Тип_наличие, соотв-х кодам типов из табл. Типы
Private Sub makeTypesFKSbutton_Click()
On Error GoTo ErrHandler
Dim rstTMC As Recordset, rstTypes As Recordset, typeID As Long, typT As String, tmcId As Long

   Set rstTMC = CurrentDb.OpenRecordset("SELECT Код_ТМЦ,КопияТип FROM Тип_Наличие ")
   rstTMC.MoveLast
   rstTMC.MoveFirst
   Do Until rstTMC.EOF
      If rstTMC.Fields(1) <> "" Then
         tmcId = rstTMC.Fields(0)
         typT = rstTMC.Fields(1)
         Set rstTypes = CurrentDb.OpenRecordset("SELECT Код_типа FROM Типы " _
                                              & "WHERE Тип ='" & typT & "'")
         rstTypes.MoveLast
         typeID = rstTypes.Fields(0)
         rstTypes.Close
         CurrentDb.Execute ("UPDATE Тип_Наличие SET Тип =" & typeID & " WHERE Код_ТМЦ = " & tmcId)
      End If
      rstTMC.MoveNext
   Loop
   
   MsgBox "completed!", , strMsgBox
   
Exit Sub
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
Exit Sub
End Sub 'makeTypesFKSbutton_Click

'обработка нажатия ОК
Private Sub okButton_Click()
Dim subDiv As String, login As String, password As String

   subdivComboBox.SetFocus
   If subdivComboBox.Text <> "" Then
      loginTextBox.SetFocus
      If loginTextBox.Text <> "" Then
         passwordTextBox.SetFocus
         If passwordTextBox.Text <> "" Then
            subdivComboBox.SetFocus
            'сохраняем филиал в локальной переменной
            subDiv = subdivComboBox.Text
            loginTextBox.SetFocus
            'сохраняем логин в локальной переменной
            login = loginTextBox.Text
            passwordTextBox.SetFocus
            'сохраняем пароль в локальной переменной
            password = passwordTextBox.Text
            'экземпляр класса Recordset
            Set rst = CurrentDb.OpenRecordset("SELECT Код,admin_login,admin_password,user_login,user_password " _
               & "FROM Структура " _
               & " WHERE Перевод LIKE '" _
               & "*" & subDiv & "*'" & "")
            'проверка уникальности выбранного филиала
            If rst.RecordCount > 1 Then
               MsgBox "В вашей структуре более одного филиала с введенным названием." _
               & " Обратитесь к администратору базы данных", vbCritical, strMsgBox
               'закрыть экземпляр класса Recordset
               rst.Close
               DoCmd.Quit acQuitSaveNone
            End If
            rst.MoveFirst
            'проверка правильности логина и пароля
            If (login = rst.Fields(1) And password = rst.Fields(2)) _
               Or (login = rst.Fields(3) And password = rst.Fields(4)) Then
               auth.setId (rst.Fields(0))
               auth.setSubdiv (subDiv)
               auth.setLogin (login)
               auth.setPassword (password)
            Else
               MsgBox "Такие логин или пароль недопустимы для введенного филиала", vbCritical, strMsgBox
               rst.Close
               Exit Sub
            End If
            'закрыть экземпляр класса Recordset
            rst.Close
            Forms("Authorization").Visible = False
            DoCmd.OpenForm "ViewСклад", acNormal, , , acFormAdd, acWindowNormal
         Else
            MsgBox "Введите пароль", vbExclamation, strMsgBox
         End If
      Else
         MsgBox "Введите логин", vbExclamation, strMsgBox
      End If
   Else
      MsgBox "Выберите филиал", vbExclamation, strMsgBox
   End If
   
   ' cleaning
   'rst.Close
   
End Sub 'okButton_Click

