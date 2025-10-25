VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim InventoryFlag As Integer, number As Integer

' объявление переменных для хранения данных журнала редактирования аттрибутов
Dim be_production_date As Date
Dim be_first_use_date As Date
Dim be_meas_param_id As Long
Dim be_exec_funct_id As Long
Dim be_notes As String
Dim be_class_id As Long
Dim be_meas_unit_id As Long
Dim be_min_limit_id As Long
Dim be_max_limit_id As Long
Dim be_accuracy_grade_id As Long
Dim be_verification_date As Date
Dim be_pov_interval_id As Long
Dim be_edit_reason As String
Dim be_id As Long

'после обновления содержимого элемента управления (выбора класса ТМЦ из выпадающего списка)
Private Sub classComboBox_AfterUpdate()

   If classComboBox.Text Like "*редства*змерени*" Then
      Call Metrological(True)
   End If
   
End Sub 'classComboBox_AfterUpdate
' нажатие кнопки Выход
Private Sub closeButton_Click()

   DoCmd.Close acForm, "EditForm"
   
End Sub ' closeButton_Click
' нажатие Редактировать
Private Sub editMetrologicalButton_Click()
   
   classComboBox.SetFocus
   If classComboBox.Text Like "*редства*змерени*" Then
      Call Metrological(True)
   End If
   
   metrologicalTab.SetFocus

End Sub ' editMetrologicalButton_Click

'обработка события происходящего сразу после вставки новой записи в БД
Private Sub Form_AfterInsert()
   
   Forms("ViewСклад").Refresh
   Call Metrological(False)
   
   
End Sub 'Form_AfterInsert

'перед обновлением данных формы
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
   Resume Next
   
End Sub 'Form_BeforeUpdate

'загрузка формы
Private Sub Form_Load()
On Error GoTo ErrHandler

   'Деактивация полей первичных аттрибутов, недоступных для редактирования
   Дата_регистрации.Enabled = False
   Зав_номер.Enabled = False
   Количество.Enabled = False
   'Деактивация полей со списками первичных аттрибутов, недоступных для редактирования
   Наименования_Код.Enabled = False
   Тип.Enabled = False
   Инв_номер_Код.Enabled = False
   Состояние.Enabled = False
   
   'деактивация элементов управления группы "Метрол. хар-ки"
   measValComboBox.Enabled = False
   unit_nameComboBox.Enabled = False
   signRuTextBox.Enabled = False
   minLimitsComboBox.Enabled = False
   maxLimitsComboBox.Enabled = False
   accuracy_grade_ComboBox.Enabled = False
   verification_date_TextBox.Enabled = False
   pov_interval_ComboBox.Enabled = False
   
   newTypeValue = Null
   InventoryFlag = 0
   
   ' вычисление строки SQL-запроса источника данных для формы редактирования аттрибутов
   sr.SelInsertСклад

   ' источником данных формы становится модифицированный запрос,
   ' который ворачивает только определенные записи из всего набора InsertСклад
   Me.RecordSource = queryString
   ' инициализация переменных для хранения данных журнала редактирования аттрибутов
   Call VarsBefore
   
   ' вызов процедуры инициализации источника данных элемента управления "Измерямая величина"
   Call MeasValsRowSourceSet

   Exit Sub
   
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Resume Next
End Sub
'при нажатии на "Esc"
Private Sub Form_Undo(Cancel As Integer)

   'вызвать процедуру деактивации элементов группы "Метролог хар-ки"
   Call Metrological(False)

End Sub 'Form_Undo(Cancel As Integer)

'обработка нажатия кнопки История перемещений ТМЦ
Private Sub historyButton_Click()
Dim ID As Long

   ID = Me.Код_ТМЦ.value

   DoCmd.OpenForm "History", acFormDS, , "Код_ТМЦ = " & ID, acFormReadOnly, acDialog

End Sub 'historyButton_Click

'справночние инвентарных номеров
Private Sub invGlossaryButton_Click()

   DoCmd.OpenForm "InvGlossary", , , , , acDialog
   Инв_номер_Код.Requery

End Sub 'invGlossaryButton_Click
'нажатие кнопки Макс. шкалы
Private Sub maxLimitsGlossaryButton_Click()

   DoCmd.OpenForm "MaxLimitsGlossary", , , , , acDialog
   Me.maxLimitsComboBox.Requery

End Sub 'maxLimitsGlossaryButton_Click
' после выбора из списка измерямых параметров
Private Sub meas_param_ComboBox_AfterUpdate()

End Sub ' meas_param_ComboBox_AfterUpdate

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
' нажатие кнопки Сохранить
Private Sub saveButton_Click()

   Me.Refresh
   
End Sub 'saveButton_Click

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
   Наименования_Код.Requery
   
End Sub 'OpenNamesGlossaryButton_Click

'обработка нажатия кнопки справочника диапазонов
Private Sub openRangesGlossaryButton_Click()

   'DoCmd.OpenForm FormName:="RangesGlossary", WindowMode:=acDialog
   'Диапазон_Код.Requery

End Sub 'OpenNamesGlossaryButton_Click
'обработка нажатия кнопки справочника типов
Private Sub openTypesGlossarybutton_Click()

   DoCmd.OpenForm FormName:="TypesGlossary", WindowMode:=acDialog
   Тип.Requery
   
End Sub 'openTypesGlossarybutton_Click

'процедура вставки в БД множества записей (по значению поля Количество)
Sub InsertNumberOfRecordsIntoDB()
On Error GoTo ErrorHandler
Dim dbs As DAO.Database
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset
Dim counter As Integer

   Set dbs = CurrentDb
   
   Set qdf = dbs.CreateQueryDef("InsertNumberOfRecordsQuery")
   qdf.SQL = " SELECT * FROM UpdateСклад "
               
   Set rst = qdf.OpenRecordset()
   
   For counter = 1 To Количество.value - 1
      With rst
         .AddNew
         '!Код_ТМЦ = auto
         !Дата_регистрации = Дата_регистрации
         !Наименования_Код = Наименования_Код
         !Тип = Тип
    '    !Диапазон = Диапазон_Код
         !Количество = 1
         !notes = Примечание_Код
         !Откуда_перемещ = 147
         !Куда_израсход_Код = 226
         !Состояние = Состояние
         !Класс_ТМЦ = classComboBox
         .Update
      End With
   Next counter

   'Return database to original.
   dbs.QueryDefs.Delete "InsertNumberOfRecordsQuery"

   rst.Close
   qdf.Close
   dbs.Close

   Set rst = Nothing
   Set qdf = Nothing
   Set dbs = Nothing

   Exit Sub

ErrorHandler:
   MsgBox "Error #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Resume Next
End Sub 'InsertNumberOfRecordsIntoDB

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
' процедура инициализации переменных для хранения данных журнала редактирования аттрибутов
Private Sub VarsBefore()

   production_dateTextBox.SetFocus
   If production_dateTextBox.Text <> "" Then
      be_production_date = CDate(production_dateTextBox.Text)
   Else
   End If

   first_use_date_TextBox.SetFocus
   If first_use_date_TextBox.Text <> "" Then
      be_first_use_date = CDate(first_use_date_TextBox.Text)
   Else
   End If
   
   be_meas_param_id = meas_param_id.value
   be_exec_funct_id = exec_funct_id.value
   Примечание_Код.SetFocus
   be_notes = Примечание_Код.Text
   be_class_id = Класс_ТМЦ.value
   be_meas_unit_id = meas_unit_id.value
   be_min_limit_id = min_limit_id.value
   be_max_limit_id = max_limit_id.value
   be_accuracy_grade_id = accuracy_grade_id.value
   
   verification_date_TextBox.SetFocus
   If verification_date_TextBox.Text <> "" Then
      be_verification_date = CDate(verification_date_TextBox.Text)
   Else
   End If
   
   be_pov_interval_id = pov_interval_id.value
   be_id = Код_ТМЦ.value
   
End Sub ' VarsBefore

' процедура инициализации источника данных элемента управления measValComboBox
Private Sub MeasValsRowSourceSet()
Dim rst As Recordset, ID As Long
   
   ' запись кода ТмЦ в переменную
   ID = Код_ТМЦ.value
   ' создаем источник данных элемента управления measValComboBox, опираясь на код ТМЦ
   Set rst = CurrentDb.OpenRecordset("SELECT * FROM MeasValsQuery WHERE Код_ТМЦ = " & ID)
   ' проверяем наличие записей источника rst (они будут только если у этого ТМЦ _
   ' имеющеется значение "Измеряемый параметр")
   If rst.BOF And rst.EOF Then
      ' если нет значения "Измеряемый параметр", то
      ' не меняем источник данных элемента управления measValComboBox
      ' что позволяет пользователю выбрать любой "Измеряемый параметр"
   Else
      ' иначе пользователь не может выбрать иной "Измеряемый параметр", кроме уже имеющегося
      measValComboBox.RowSource = "SELECT * FROM MeasValsQuery WHERE Код_ТМЦ = " & ID
   End If

End Sub ' MeasValsRowSourceSet
