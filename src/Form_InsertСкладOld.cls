VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InsertСкладOld"
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
' нажатие кн. Выход
Private Sub closeButton_Click()

   DoCmd.Close acForm, "InsertСклад"
   
End Sub ' closeButton_Click

'обработка события происходящего сразу после вставки новой записи в БД
Private Sub Form_AfterInsert()
   
   Forms("InsertСклад").Requery
   Call Metrological(False)
   
End Sub 'Form_AfterInsert

'перед обновлением данных формы
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrHandler
Dim Response As Integer, TypeText As String, SerNumText As String, InvNumText As String
Dim counter As Integer
   
   Тип.SetFocus
   TypeText = Тип.Text
   Зав_номер.SetFocus
   SerNumText = Зав_номер.Text
   Инв_номер_Код.SetFocus
   InvNumText = Инв_номер_Код.Text
   
   If (TypeText <> "") And (SerNumText = "" Or InvNumText = "") Then
      Response = MsgBox("Не указаны заводской и(или) инвентарный номера! Сохранить?", _
         vbOKCancel + vbExclamation + vbDefaultButton2, strMsgBox)
      If Response = 2 Then
         Cancel = True
      End If
   End If

   Количество.SetFocus
   If Val(Количество.Text) > 1 Then
      If (SerNumText <> "" Or InvNumText <> "") Then
         Cancel = True
         MsgBox "Ввод множественного количества ТМЦ, имеющих инвентарный и(или) заводской " & _
            "номера недопустим. Такие ТМЦ можно вводить только по одному!", , strMsgBox
      Else
         'вызвать процедуру множественной вставки в БД
         Call InsertNumberOfRecordsIntoDB
         'подготовка поля Количество к завершению выполнения процедуры вставки
         'т.к. при вызове InsertNumberOfRecordsIntoDB (Количество-1) ТМЦ были вставлены в БД
         Количество.SetFocus
         Количество.Text = "1"
      End If
   End If
   
Exit Sub

ErrHandler:
   MsgBox "Error #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Me.Undo
   Resume Next
   
End Sub 'Form_BeforeUpdate
'загрузка формы
Private Sub Form_Load()
On Error GoTo ErrHandler

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

   Exit Sub
   
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Resume Next
End Sub
' перед открытием формы (InsertСклад) запрашиваем информацию о сопроводительной бумаге (бумагах)
' Как правило - это стандартные требования-накладные сопровождающие ТМЦ
Private Sub Form_Open(Cancel As Integer)

    ' запуск формы получения информации о сопроводительных бумагах
    
End Sub

'при нажатии на "Esc"
Private Sub Form_Undo(Cancel As Integer)

   'вызвать процедуру деактивации элементов группы "Метролог хар-ки"
   Call Metrological(False)

End Sub 'Form_Undo(Cancel As Integer)

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
' нажатие кн. Сохранить
Private Sub saveButton_Click()

   Me.Refresh
   
End Sub ' saveButton_Click

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
