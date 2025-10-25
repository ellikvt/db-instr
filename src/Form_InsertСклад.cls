VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InsertСклад"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim InventoryFlag As Integer, number As Integer
'после обновления содержимого элемента управления (выбора класса ТМЦ из выпадающего списка)
Private Sub classComboBox_AfterUpdate()

   If classComboBox.Text Like "*редства*змерени*" Then
      Call Metrological(True)
   End If
   
End Sub 'classComboBox_AfterUpdate

Private Sub classComboBox_BeforeUpdate(Cancel As Integer)
 
 ' Check for a blank value in the classComboBox field.
    If IsNull(Me.classComboBox) Then

       ' Alert the user.
       MsgBox "Вы должны указать Класс ТМЦ", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

End Sub ' classComboBox_BeforeUpdate

' нажатие кн. Выход
Private Sub closeButton_Click()
   
    DoCmd.Close acForm, "InsertСклад"
   
End Sub ' closeButton_Click

Private Sub defaultModeButton_Click()
    
    Me.typesModeButton.Caption = "Выключен"
    Call namesComboBox_AfterUpdate
    
End Sub

'загрузка формы
Private Sub Form_Load()
On Error GoTo ErrHandler

    Call clearForm
    Me.typesModeButton.Caption = "Выключен"
    Me.defaultModeButton.Caption = "По умолчанию"
    Exit Sub
   
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If MsgBox("Вы уверены, что хотите закрыть эту форму?", vbYesNo) = vbYes Then
        Exit Sub
    Else
        Cancel = True
    End If

End Sub

'справочник инвентарных номеров
Private Sub invGlossaryButton_Click()

   DoCmd.OpenForm "InvGlossary", , , , , acDialog
   Me.inventoryComboBox.Requery

End Sub 'invGlossaryButton_Click
'нажатие кнопки Макс. шкалы
Private Sub maxLimitsGlossaryButton_Click()

   DoCmd.OpenForm "MaxLimitsGlossary", , , , , acDialog
   Me.maxLimitsComboBox.Requery

End Sub 'maxLimitsGlossaryButton_Click

'после выбора из списка измерямых величин
Private Sub measValComboBox_AfterUpdate()
   
   'обновить источник строк списка единиц измерения unit_nameComboBox
   'в зависимости от выбранной в списке measValComboBox измеряемой величины
   Me.unit_nameComboBox.RowSource = "SELECT unit_name, meas_unit_id FROM" & _
                            " meas_units WHERE meas_val_id = " & Me.measValComboBox
                                                        

End Sub 'measValComboBox_AfterUpdate
'нажатие кнопки Мин. шкалы
Private Sub minLimitsGlossaryButton_Click()

   DoCmd.OpenForm "MinLimitsGlossary", , , , , acDialog
   Me.minLimitsComboBox.Requery
   
End Sub 'minLimitsGlossaryButton_Click
' просмотр списка имен
Private Sub namesComboBox_AfterUpdate()
    
    If Me.typesModeButton.Caption = "Выключен" Then
    ' Update the row source of the types combo box
    ' when the user makes a selection in the names
    ' combo box.
        Me.typesComboBox.RowSource = "SELECT Тип " & _
            "FROM Типы INNER JOIN (Наименования INNER JOIN Тип_Наличие ON Наименования.Код = Тип_Наличие.Наименования_Код) " & _
                "ON Типы.Код_типа = Тип_Наличие.Тип_Код " & _
            "GROUP BY Типы.Тип, Наименования.F4 " & _
            "HAVING (((Наименования.F4) = """ & Me.namesComboBox & """)) " & _
            "ORDER BY Тип"
    End If
                            
    Me.typesComboBox = Me.typesComboBox.ItemData(0)

End Sub ' namesComboBox_AfterUpdate()

Private Sub namesComboBox_BeforeUpdate(Cancel As Integer)

 ' Check for a blank value in the namesComboBox field.
    If IsNull(Me.namesComboBox) Then

       ' Alert the user.
       MsgBox "Вы должны указать Наименование", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

End Sub ' namesComboBox_BeforeUpdate(Cancel As Integer)

Private Sub namesComboBox_NotInList(NewData As String, Response As Integer)
    
    If MsgBox("Нет в списке. Изменить элементы списка?", vbYesNo + vbQuestion, strMsgBox) = vbYes Then
        DoCmd.OpenForm FormName:="NamesGlossary", WindowMode:=acDialog
        Me.namesComboBox = ""
        Me.namesComboBox.Requery
    End If
    
End Sub

Private Sub numberTextBox_BeforeUpdate(Cancel As Integer)

 ' Check for a blank value in the numberTextBox field.
    If IsNull(Me.numberTextBox) Then

       ' Alert the user.
       MsgBox "Вы должны указать Количество", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

 ' Check for a greater then zero value in the numberTextBox field.
    If Me.numberTextBox < 0 Or Me.numberTextBox = 0 Then

       ' Alert the user.
       MsgBox "Вы должны указать Количество >= 0", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

End Sub ' numberTextBox_BeforeUpdate

Private Sub registrationDateTextBox_BeforeUpdate(Cancel As Integer)
 
 ' Check for a blank value in the registrationDateTextBoxt field.
    If IsNull(Me.registrationDateTextBox) Then

       ' Alert the user.
       MsgBox "Вы должны указать дату ввода в базу данных", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

End Sub 'registrationDateTextBox_BeforeUpdate

' нажатие кн. Сохранить
Private Sub saveButton_Click()
On Error GoTo ErrorHandler
' обязат-е вспомогат. переменные процедуры вставки для сбора данных
' о типе, инв. номера, наименовании ТМЦ
Dim typeID As Variant, invNumID As Variant, nameID As Long, stateID As Long
' для преобразования наименований, имеющих знаки кавычек
Dim tname As String
' для извлечения и хранения значений из таблиц БД
Dim hrst As Recordset
Dim counter As Integer ' счетчик количества ТМЦ

    If Me.namesComboBox = "" Then ' если Наименование пропущено
        MsgBox "Не указано значение в поле Наименование", vbExclamation + vbOKOnly, strMsgBox
        Me.namesComboBox.SetFocus
        Exit Sub
    End If
    If Me.numberTextBox = "" Then ' если Количество пропущено
        MsgBox "Не указано значение в поле Количество", vbExclamation + vbOKOnly, strMsgBox
        Me.numberTextBox.SetFocus
        Exit Sub
    End If
    If Me.statesComboBox <> "" Then ' если Состояние пропущено
    Else
        MsgBox "Не указано значение в поле Состояние", vbExclamation + vbOKOnly, strMsgBox
        Me.statesComboBox.SetFocus
        Exit Sub
    End If
    If Me.classComboBox = "" Then ' если класс ТМЦ пропущен
        MsgBox "Не указано значение в поле Класс ТМЦ", vbExclamation + vbOKOnly, strMsgBox
        Me.classComboBox.SetFocus
        Exit Sub
    End If
    
    ' Блок 1 обработки условий при вычислении значения кода Наименования ТМЦ
    If Me.namesComboBox Like "*""*" Then ' если наименование содержит кавычку (-и)
        ' заменить кавычку (-и) двойной (-ными)
        tname = Replace(Me.namesComboBox, """", """""")
        ' запрос кода Наименования ТМЦ
        Set hrst = CurrentDb.OpenRecordset("SELECT Наименования.Код FROM Наименования WHERE F4= """ & tname & """")
    Else
    ' запрос кода Наименования ТМЦ
        Set hrst = CurrentDb.OpenRecordset("SELECT Наименования.Код FROM Наименования WHERE F4= """ & Me.namesComboBox & """")
    End If
    ' Блок 2 обработки условий при вычислении значения кода Наименования ТМЦ
    If Not (hrst.BOF And hrst.EOF) Then
      nameID = hrst.Fields(0)
    Else
      MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
      Me.namesComboBox.SetFocus
      Exit Sub
    End If
    
    ' Блок обработки условий при вычислении значения кода Типа ТМЦ
    If Me.typesComboBox Like "*""*" Then ' если тип содержит кавычку (-и)
        ' заменить кавычку (-и) двойной (-ными)
        tname = Replace(Me.typesComboBox, """", """""")
        ' запрос кода Типа
        Set hrst = CurrentDb.OpenRecordset("SELECT Типы.Код_типа FROM Типы WHERE Тип= """ & tname & """")
        If Not (hrst.BOF And hrst.EOF) Then
            typeID = hrst.Fields(0)
        Else
            MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
            Me.typesComboBox.SetFocus
            Exit Sub
        End If
    Else
        If Me.typesComboBox <> "" Then
            ' запрос кода Типа
            Set hrst = CurrentDb.OpenRecordset("SELECT Типы.Код_типа FROM Типы WHERE Тип= """ & Me.typesComboBox & """")
            If Not (hrst.BOF And hrst.EOF) Then
                typeID = hrst.Fields(0)
            Else
                MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
                Me.typesComboBox.SetFocus
                Exit Sub
            End If
        Else
            ' Типа не существует
            typeID = Null
        End If
    End If
    
    ' Блок обработки условий при вычислении значения кода Инв. номера ТМЦ
    If Me.inventoryComboBox <> "" Then ' если инв. номер введен
        Set hrst = CurrentDb.OpenRecordset("SELECT Инв_номер.Код FROM Инв_номер WHERE F6= """ & Me.inventoryComboBox & """")
        If Not (hrst.BOF And hrst.EOF) Then
            invNumID = hrst.Fields(0)
        Else
            MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
            Me.inventoryComboBox.SetFocus
            Exit Sub
        End If
    Else
        ' Инвентарного номера не существует
        invNumID = Null
    End If
    
    ' Блок вычисления значения кода Состояния ТМЦ
    Set hrst = CurrentDb.OpenRecordset("SELECT state_id FROM states WHERE state = """ & Me.statesComboBox & """")
    If Not (hrst.BOF And hrst.EOF) Then
        stateID = hrst.Fields(0)
    Else
        MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
        Me.statesComboBox.SetFocus
        Exit Sub
    End If
    
            
    '*********************************************************************************************
    '          Проверки непротиворечивости данных, после того как обязательные данные введены
    '*********************************************************************************************
    If (Me.typesComboBox <> "") And (Me.serNumTextBox = "" Or Me.inventoryComboBox = "") Then
        Response = MsgBox("Не указаны заводской и(или) инвентарный номера! Сохранить?", _
         vbOKCancel + vbExclamation + vbDefaultButton2, strMsgBox)
        If Response = 2 Then
            If Me.serNumTextBox = "" Then
                Me.serNumTextBox.SetFocus
            Else
                Me.inventoryComboBox.SetFocus
            End If
            Exit Sub
        End If
    End If

    If Me.numberTextBox > 1 Then
        If (Me.serNumTextBox <> "" Or Me.inventoryComboBox <> "") Then
            MsgBox "Ввод множественного количества ТМЦ, имеющих инвентарный и(или) заводской " & _
            "номера недопустим. Такие ТМЦ можно вводить только по одному!", vbExclamation, strMsgBox
            Me.numberTextBox.SetFocus
            Exit Sub
        End If
    End If
    '**********************************************************************************************
    counter = Me.numberTextBox
    ' вставка ТМЦ в БД
    Do
        CurrentDb.Execute "INSERT INTO Тип_Наличие (Дата_регистрации, Откуда_перемещ, Тип_Код, Зав_номер, Количество," & _
                           "Инв_номер_Код, Куда_израсход_Код, Наименования_Код, notes, lookup_state," & _
                           "class_id, verification_date, accuracy_grade_id, max_limit_id, meas_unit_id," & _
                           "min_limit_id, pov_interval_id, production_date, first_use_date) " & _
                           "VALUES ('" & Me.registrationDateTextBox & "','" & SOURCE_ID & "','" & typeID & "','" & Me.serNumTextBox & "', '1','" & _
                           invNumID & "','" & DESTINATION_ID & "','" & nameID & "','" & Me.notesTextBox & "','" & stateID & "','" & _
                           Me.classComboBox & "','" & Me.verification_date_TextBox & "','" & Me.accuracy_grade_ComboBox & "','" & _
                           Me.maxLimitsComboBox & "','" & Me.unit_nameComboBox & "','" & Me.minLimitsComboBox & "','" & _
                           Me.pov_interval_ComboBox & "','" & Me.production_dateTextBox & "','" & Me.first_use_dateTextBox & "')"
        counter = counter - 1
    Loop While counter > 0 ' организация цикла с постпроверкой условия ( так как ТМЦ могут быть кол-вом больше 1 )
    MsgBox "Транзакция проведена успешно!", vbInformation, strMsgBox
    Call clearForm
    Exit Sub

ErrorHandler:
   MsgBox "Error #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Resume Next
End Sub ' saveButton_Click

Private Sub statesComboBox_BeforeUpdate(Cancel As Integer)
 
 ' Check for a blank value in the statesComboBox field.
    If IsNull(Me.statesComboBox) Then

       ' Alert the user.
       MsgBox "Вы должны указать Состояние", vbExclamation, strMsgBox

      ' Cancel the update.
      Cancel = True
    End If

End Sub ' statesComboBox_BeforeUpdate

Private Sub typesModeButton_Click()
    
    If Me.typesModeButton.Caption = "Выключен" Then
        Me.typesModeButton.Caption = "Включен"
        Me.typesComboBox.RowSource = "SELECT Тип FROM Типы ORDER BY Тип"
    Else
        Me.typesModeButton.Caption = "Выключен"
        Call namesComboBox_AfterUpdate
    End If
    
 
End Sub ' typesModeButton_Click

'после выбора из списка единиц измерения
Private Sub unit_nameComboBox_AfterUpdate()
Dim rst As Recordset

   'выбрать соответствующее русское обозначение единицы измерения
   Set rst = CurrentDb.OpenRecordset("SELECT sign_ru FROM" & _
                            " meas_units WHERE meas_unit_id = " & Me.unit_nameComboBox)
   If Not (rst.BOF And rst.EOF) Then
      'вывести русское обозначение в соответствующем поле
      signRuTextBox.SetFocus
      signRuTextBox = rst.Fields(0)
   Else
      MsgBox "Нет соответствующей единицы в базе. Обратитесь к администратору базы данных", vbExclamation + vbOKOnly, strMsgBox
   End If

End Sub 'unit_nameComboBox_AfterUpdate

'обработка нажатия кнопки справочника наименований
Private Sub OpenNamesGlossaryButton_Click()

   DoCmd.OpenForm FormName:="NamesGlossary", WindowMode:=acDialog
    Me.namesComboBox.Requery
   
End Sub 'OpenNamesGlossaryButton_Click

'обработка нажатия кнопки справочника типов
Private Sub openTypesGlossarybutton_Click()

   DoCmd.OpenForm FormName:="TypesGlossary", WindowMode:=acDialog
   Me.typesComboBox.Requery
   
End Sub 'openTypesGlossarybutton_Click

'процедура управлением активации/деактивации элементов группы "Метрол. хар-ки"
Private Sub Metrological(enab As Boolean)

   If enab Then
      'активация элементов управления группы "Метрол. хар-ки"
      measValComboBox.Enabled = True
      unit_nameComboBox.Enabled = True
      signRuTextBox.Enabled = True
      minLimitsComboBox.Enabled = True
      maxLimitsComboBox.Enabled = True
      accuracy_grade_ComboBox.Enabled = True
      verification_date_TextBox.Enabled = True
      pov_interval_ComboBox.Enabled = True
   Else
      production_dateTextBox.SetFocus
      'деактивация элементов управления группы "Метрол. хар-ки"
      measValComboBox.Enabled = False
      unit_nameComboBox.Enabled = False
      signRuTextBox.Enabled = False
      minLimitsComboBox.Enabled = False
      maxLimitsComboBox.Enabled = False
      accuracy_grade_ComboBox.Enabled = False
      verification_date_TextBox.Enabled = False
      pov_interval_ComboBox.Enabled = False
   End If

End Sub

' подготовка формы к вводу следующего ТМЦ
Private Sub clearForm()

    Call Metrological(False)
   newTypeValue = Null
   InventoryFlag = 0
    
    Me.namesComboBox.RowSource = "SELECT F4 FROM Наименования ORDER BY F4"
    Me.namesComboBox = ""
    Me.typesComboBox.RowSource = "SELECT Тип FROM Типы ORDER BY Тип"
    Me.typesComboBox = ""
    Me.inventoryComboBox.RowSource = "SELECT F6 FROM Инв_номер ORDER BY F6"
    Me.inventoryComboBox = ""
    Me.statesComboBox.RowSource = "SELECT state FROM states ORDER BY state"
    Me.statesComboBox = ""
    Me.classComboBox = ""
    Me.production_dateTextBox = ""
    Me.first_use_dateTextBox = ""
    Me.serNumTextBox = ""
    Me.numberTextBox = ""
    Me.notesTextBox = ""
    
End Sub ' clearForm()

