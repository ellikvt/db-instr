VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_UpdateСклад"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim targetRelocChanged As Boolean
'объект, которому будет присвоен экземпляр класса clsTreeClass
Dim tr As Object
' переменная для хранения количества перемещенных ТМЦ (обнуляется после успешного перемещения)
Dim rlNumber As Integer
' переменная для хранения назначения строки-места назначения перемещаемого ТМЦ (обнуляется после успешного перемещения)
Dim rlStr As String
Dim clickCounter As Integer
'объявление переменных хранения данных для исторической записи
Dim rDate As Date 'дата перемещения
Dim bRplace_id As Long 'код местонахождения до перемещения
Dim brPersonal As String 'ответственный до перемещения
Dim bRstate_id As Long 'код состояния до перемещения
Dim bCclass_id As Long 'код класса ТМЦ до его смены
Dim bRsession_id As Long 'код текущей сессии
Dim bRreason_id As Long 'код причины до перемещения
Dim someStr As String
' нажатие кн. Выйти
Private Sub closeButton_Click()

   DoCmd.Close acForm, "UpdateСклад"
   
End Sub ' closeButton_Click

'обработка действий после обновления записи
Private Sub Form_AfterUpdate()
Dim rst As Recordset

   Forms("UpdateСклад").Requery
   Set rst = Me.Recordset
   
   If rst.RecordCount = 0 Then 'если нет больше записей, т.е. все ТМЦ перемещены
      MsgBox "Успешно перемещено " & rlNumber & " шт. ТМЦ данного типа и наименования" & vbCrLf _
           & "по назначению: " & rlStr, vbInformation + vbOKOnly, strMsgBox
      MsgBox "Все заданные перемещения обработаны!" & vbCrLf _
           & "Форма перещений будет закрыта.", vbInformation + vbOKOnly, strMsgBox
      DoCmd.Close acForm, "UpdateСклад", acSaveNo
      rst.Close
   Else                          'если не все ТМЦ перемещены
      'снова блокируем поле Куда от случайных изменений
      Куда_израсход_Код.Enabled = False
      'populate the Recordset
      rst.MoveLast
      'обновить глобальную переменную, содержащую количество записей формы
      rl.recsCount = rst.RecordCount
      'снова очистить поле "Количество" в группе "Сведения о перемещении"
      relocNumberTextBox.value = ""
      ' активируем кнопку "Начало перемещения"
      Me.relocateButton.Enabled = True
      MsgBox "Успешно перемещено " & rlNumber & " шт. ТМЦ данного типа и наименования" & vbCrLf _
           & "по назначению: " & rlStr, vbInformation + vbOKOnly, strMsgBox
   End If
   
   ' сброс значения переменной для хранения количества перемещенных ТМЦ в нуль
   rlNumber = 0
   ' сброс значения переменной для хранения строки - места назначения перемещаемого ТМЦ
   rlStr = ""
   
Exit Sub
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description & " Form_AfterUpdate()"
   Resume Next
    
End Sub 'Form_AfterUpdate
'выполнение проверки перед обновлением записи
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo ErrHandler
Dim Result As Integer, ID As Long
Dim rst As Recordset

    ' условие что это перемещение не типа "структура - буфер"
    If Not (rl.insOutRelocation) Then
        'проверка свойства, что назначение - это "позиция"
        If rl.isItPosition(hidedTextBox.value) Then
            'проверка занятости позиции прибором
            If rl.busyPosition Then
                MsgBox "Позиция занята", , strMsgBox
                Cancel = True
                Exit Sub
            End If
        End If
       
        'проверка свойства, что назначение - это "место хранения"
        If rl.isItStore(hidedTextBox.value) Then
        End If
    
        'проверка свойства, что назначение - это "использованные ТМЦ"
        If rl.isItUsed_Store(hidedTextBox.value) Then
        End If
       
        'проверка свойства, что назначение - это не "место хранения", "использованные ТМЦ" или "позиция"
        If Not (rl.toPosition Or rl.toStore Or rl.toUsed_Store) Then
            MsgBox "Не позиция, не склад и не использованные", , strMsgBox
            Cancel = True
            Exit Sub
        End If
        
        If Not rl.outRelocation Then
            ID = Код_ТМЦ.value
            Set rst = CurrentDb.OpenRecordset("SELECT structure_id FROM " _
                                            & "Тип_Наличие WHERE " _
                                            & "Тип_Наличие.Код_ТМЦ = " & ID & ";")
            rst.MoveFirst
            If auth.getId <> getDesiredSubdivCode(rst.Fields(0)) Then
                ' закрыть источник данных
                rst.Close
                MsgBox "При авторизации в базе данных вами указан другой филиал." & vbCrLf _
                   & "Возможная причина появления этого сообщения:" & vbCrLf _
                   & "1. Вы хотите переместить ТМЦ из другого филиала" & vbCrLf _
                   & "2. Вы хотите переместить ТМЦ в другой филиал" & vbCrLf _
                   & "Перемещение ТМЦ между филиалами производятся " & vbCrLf _
                   & "только через буфер обмена приборов" & vbCrLf _
                   & "Операция будет отменена", vbExclamation + vbOKOnly, strMsgBox
                 Cancel = True
                 Me.Undo
                 Exit Sub
            End If
        End If
            
        'проверка условия, что выбранный филиал не тот, под которым пользователь авторизован
        If auth.getId <> getDesiredSubdivCode(hidedTextBox.value) Then
           MsgBox "При авторизации в базе данных вами указан другой филиал." & vbCrLf _
              & "Возможная причина появления этого сообщения:" & vbCrLf _
              & "1. Вы хотите переместить ТМЦ из другого филиала" & vbCrLf _
              & "2. Вы хотите переместить ТМЦ в другой филиал" & vbCrLf _
              & "Перемещение ТМЦ между филиалами производятся " & vbCrLf _
              & "только через буфер обмена приборов" & vbCrLf _
              & "Операция будет отменена", vbExclamation + vbOKOnly, strMsgBox
            Cancel = True
            Me.Undo
            Exit Sub
        End If
    Else ' если это перемещение типа "структура - буфер"
        structure_idTextBox.SetFocus
        structure_idTextBox.Text = ""
        targetRelocChanged = True
    End If
      
   If (targetRelocChanged = False) Or (Куда_израсход_Код.Enabled = False) Then
      Cancel = True
      MsgBox "Не менялось поле 'Куда перемещено'", vbCritical, strMsgBox
   End If
   
   If Cancel = False Then
      Result = MsgBox("Производится запись заданных аттрибутов в базу. Отмена изменений будет невозможна. Продолжить?", vbInformation + vbOKCancel + vbDefaultButton1, strMsgBox)
      If Result = 2 Then
         Cancel = True
         Me.Undo
         Exit Sub
         'fromPlaceLabel.Caption = "Откуда перемещено"
         '***********************************************************************
         'relocationDateTextBox.SetFocus 'отвлечение внимания от поля Куда израсх
                                          'для измения св-в этого эл-та управления
         'If Куда_израсход_Код.Locked = False Then
         '   Куда_израсход_Код.Locked = True
         'End If
         
         'If Куда_израсход_Код.Enabled = True Then
         '   Куда_израсход_Код.Enabled = False
         'End If
         '***********************************************************************
      End If
   End If
   
   'вставка строки в таблицу history_records, используя значения _
   из переменных хранения данных для исторической записи
   CurrentDb.Execute " INSERT INTO history_records(relocation_date,Куда_израсход_Код, " _
                     & "Инициалы,state_id,class_id,session_id,reason_id,Код_ТМЦ ) VALUES " _
                     & "('" & rDate & "','" _
                     & bRplace_id & "','" _
                     & brPersonal & "','" _
                     & bRstate_id & "','" _
                     & bCclass_id & "','" _
                     & bRsession_id & "','" _
                     & bRreason_id & "','" _
                     & Код_ТМЦ & " ' );"
   
   relocNumberTextBox.SetFocus
   If Val(relocNumberTextBox.Text) = 1 Then
      rlNumber = 1
   Else
      'вызвать процедуру множественной вставки в БД bRreason_id
      Call UpdateRecords
      'подготовка поля Количество к завершению выполнения процедуры вставки
      'т.к. при вызове UpdateRecords (Количество-1) ТМЦ были вставлены в БД
      'relocNumberTextBox.SetFocus
      'relocNumberTextBox.Locked = False
      'relocNumberTextBox.Text = "1"
      'relocNumberTextBox.Locked = True
      
   End If
      
Exit Sub
ErrHandler:
   MsgBox "Error #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   Me.Undo
   Resume Next

End Sub 'Form_BeforeUpdate

'загрузка формы Перемещение
Private Sub Form_Load()
On Error GoTo ErrHandler
   
   Me.TimerInterval = 1000
   If Not (rl.insOutRelocation) Then
      'инициализация метки подсказки
      TreeViewNotesLabel.Caption = "* Подсказка: нажмите 'Начало перемещения' " _
         & " В открывшейся форме заполните все поля." _
         & " Затем можно будет выбрать назначение в дереве структуры."
   
      'Инициализация дерева структуры
      Set tr = New clsTreeClass
      Set tr.Tree = Me.TrView.Object
      tr.Tbl = "Структура"
      tr.fldKey = "Код"
      tr.fldParent = "parent_id"
      tr.fldText = "Наименование"
      tr.GenerateTree
   Else
      TrView.Enabled = False
      'инициализация метки подсказки
      TreeViewNotesLabel.Caption = "* Подсказка: нажмите 'Начало перемещения' " _
         & " В открывшейся форме заполните все поля." _
         & " Затем проведите запись информации о перемещении в базу данных."
   End If
   
   Me.RecordSource = queryString
   'создаем новый объект Recordset, содержащий все записи запускаемой формы
   ' инициализация переменной для хранения количества перемещенных ТМЦ (обнуляется после успешного перемещения)
   rlNumber = 0
         
Exit Sub
ErrHandler:
   MsgBox Err.Description
End Sub 'Form_Load
'процедура обработки события по таймеру - реализована простая анимация метки подсказки
Private Sub Form_Timer()
Static ShowLabel As Boolean, counter As Integer
    
    If counter > 10 Then
      Me.TimerInterval = 0
    Else
      If ShowLabel Then
          ' показать метку подсказки.
          TreeViewNotesLabel.BackColor = -2147483623
      Else
          ' Не показывать метку подсказки.
          TreeViewNotesLabel.BackColor = -2147483624
      End If
      ShowLabel = Not ShowLabel
      counter = counter + 1
   End If
 
End Sub 'Form_Timer

'обработка нажатия "Esc"
Private Sub Form_Undo(Cancel As Integer)
   
   cancelUpdateQwr = False
   relocNumberTextBox.SetFocus
   relocNumberTextBox.Locked = False
   relocNumberTextBox.Text = ""
   relocNumberTextBox.Locked = True
   relocationDateTextBox.SetFocus 'отвлечение внимания от поля Куда израсх _
                                    для измения св-в этого эл-та управления
   If Куда_израсход_Код.Locked = False Then
      Куда_израсход_Код.Locked = True
   End If
   
   If Куда_израсход_Код.Enabled = True Then
      Куда_израсход_Код.Enabled = False
   End If
   
End Sub 'Form_Undo

'обработка нажатия кнопки справочника примечаний
Private Sub openNotesGlossarybutton_Click()

   DoCmd.OpenForm FormName:="NotesGlossary", WindowMode:=acDialog
   
   Примечание_Код.Requery
   Me.Requery

End Sub 'openNotesGlossarybutton_Click

'обработка нажатия кнопки справочника применений ТМЦ
Private Sub OpenUseGlossary_Click()

   DoCmd.OpenForm FormName:="UseGlossary", WindowMode:=acDialog
   
   Куда_израсход_Код.Requery
   Me.Requery

End Sub 'OpenUseGlossary_Click

Private Sub Form_Unload(Cancel As Integer)

   'установка свойст класса Перемещение в исходные значения
   rl.outRelocation = False
   rl.inRelocation = False
   rl.insOutRelocation = False
   rl.toPosition = False
   rl.toStore = False
   
End Sub

'нажатие кнопки "Начало перемещения.."
Private Sub relocateButton_Click()
On Error GoTo ErrHandler
Dim rs As Recordset
Dim rsS As Recordset 'для определения кодов переменных (из таблиц classes, states, Куда_израсход, sessions), _
                      входящих в историческую запись таблицы history_records
Dim bookmark As Variant
Dim str As String ' строковая переменная для копирования значения поля Местонахождение
                  ' в поле Откуда перемещено
   
   'выбор данных для истор. записи, запоминание данных для истор. записи в переменных
   relocationDateTextBox.SetFocus
   relocationDateTextBox.Locked = False
   If relocationDateTextBox.Text <> "" Then
      rDate = CDate(relocationDateTextBox.Text)
   Else
   End If
   relocationDateTextBox.Locked = True
   
   placeTextBox.SetFocus
   placeTextBox.Locked = False
   someStr = placeTextBox.Text
   placeTextBox.Locked = True
   'запрос bRplace_id
   Set rsS = CurrentDb.OpenRecordset("SELECT Код FROM Куда_израсход " _
                                   & "WHERE F27 = '" & someStr & "'")
   If rsS.RecordCount = 1 Then
      bRplace_id = rsS.Fields(0)
   Else
      MsgBox "Текущее Местонахождение отсутствует или неуникально. Ошибка запроса к таблице." _
             & " Обратитесь к администратору базы данных.", vbExclamation, strMsgBox
      rsS.Close
      Exit Sub
   End If
   rsS.Close
      
   personalTextBox.SetFocus
   personalTextBox.Locked = False
   brPersonal = personalTextBox.Text
   personalTextBox.Locked = True
   
   stateComboBox.SetFocus
   stateComboBox.Locked = False
   someStr = stateComboBox.Text
   stateComboBox.Locked = True
   'запрос bRstate_id
   Set rsS = CurrentDb.OpenRecordset("SELECT state_id FROM states " _
                                   & "WHERE state = '" & someStr & "'")
   If rsS.RecordCount = 1 Then
      bRstate_id = rsS.Fields(0)
   Else
      MsgBox "Текущее Состояние отсутствует или неуникально. Ошибка запроса к таблице." _
             & " Обратитесь к администратору базы данных.", vbExclamation, strMsgBox
      rsS.Close
      Exit Sub
   End If
   rsS.Close
     
   classTextBox.SetFocus
   classTextBox.Locked = False
   someStr = classTextBox.Text
   classTextBox.Locked = True
   'запрос bCclass_id
   Set rsS = CurrentDb.OpenRecordset("SELECT class_id FROM classes " _
                                   & "WHERE class_name = '" & someStr & "'")
   If rsS.RecordCount = 1 Then
      bCclass_id = rsS.Fields(0)
   Else
      MsgBox "Текущий класс ТМЦ отсутствует или неуникален. Ошибка запроса к таблице." _
             & " Обратитесь к администратору базы данных.", vbExclamation, strMsgBox
      rsS.Close
      Exit Sub
   End If
   rsS.Close
   
   reasonComboBox.SetFocus
   reasonComboBox.Locked = False
   someStr = reasonComboBox.Text
   reasonComboBox.Locked = True
   'запрос bRreason_id
   Set rsS = CurrentDb.OpenRecordset("SELECT reason_id FROM reasons " _
                                   & "WHERE reason = '" & someStr & "'")
   If rsS.RecordCount = 1 Then
      bRreason_id = rsS.Fields(0)
   Else
      MsgBox "Текущее основание перемещения отсутствует или неуникально. Ошибка запроса к таблице." _
             & " Обратитесь к администратору базы данных.", vbExclamation, strMsgBox
      rsS.Close
      Exit Sub
   End If
   rsS.Close
                                   
   Set rsS = CurrentDb.OpenRecordset("SELECT session_id" _
                                   & " FROM sessions " _
                                   & " WHERE subdivision Like '" _
                                   & auth.getSubdiv & "' " _
                                   & " AND user_name Like '" _
                                   & auth.getLogin & "' " _
                                   & " AND state Like " _
                                   & "'открыта'" _
                                   & " AND host_name Like '*" _
                                   & GetThisComputerName & "*'" _
                                   & " ORDER BY sessions.session_id")
   rsS.MoveLast
   bRsession_id = rsS.Fields(0)
   rsS.Close
      
   Set rs = Me.Recordset
   bookmark = rs.bookmark
   'populate the Recordset
   rs.MoveLast
   'инициализируем счетчик количества записей формы - глобальную переменную
   rl.recsCount = rs.RecordCount
   rs.bookmark = bookmark
   
   DoCmd.OpenForm "RelocateForm", WindowMode:=acDialog
   
   'если предыдущие данные с формы Сбора параметров о перемещении получены все
   'тогда разблокируем поле ввода места назначения "Куда перемещено"
   If cancelUpdateQwr = False Then
      Куда_израсход_Код.Enabled = True
      Куда_израсход_Код.Locked = True
      If Not rl.insOutRelocation Then
      Else
         Me.placeTextBox.SetFocus
         str = Me.placeTextBox.Text
         Me.Местонахождение.SetFocus
         Me.Местонахождение.Locked = False
         Me.Местонахождение.Text = str
         Me.Местонахождение.Locked = True
         Me.Куда_израсход_Код.SetFocus
         Куда_израсход_Код.Locked = False
         Куда_израсход_Код.Text = "буфер приборов"
         rlStr = "буфер приборов"
         Куда_израсход_Код.Locked = True
      End If
      Куда_израсход_Код.SetFocus

      ' деактивируем кнопку "Начало перемещения"
      Me.relocateButton.Enabled = False
   Else
      Me.Undo
      'если идет одновременная работа, то число записей меняется и нужно обновлять объект
      Me.Requery
   End If
   
Exit Sub
ErrHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description & " в процедуре Private Sub relocateButton_Click формы UpdateСклад"
   Resume Next
End Sub 'relocateButton_Click
' нажатие кн. Сохранить
Private Sub saveButton_Click()

   Me.Refresh
   
End Sub ' saveButton_Click

'обработка кликов мышкой на дереве структуры
Private Sub TrView_Click()
On Error GoTo ErrHandler
Dim rs As Recordset
Dim i As Integer, ID As Long, pathStr As String, oldPathStr As String
Dim pathArray(14) As String

   'запись индекса выбранного узла в скрытом поле на форме
   hidedTextBox.value = tr.GetKey
   ID = hidedTextBox.value
      
   i = 0
   Call arrayInitilize(pathArray)
   Do
      Set rs = CurrentDb.OpenRecordset("SELECT Перевод,parent_id FROM Структура " _
           & " WHERE Код = " _
           & ID & "")
      ID = rs("parent_id")
      pathArray(i) = rs("Перевод")
      i = i + 1
   Loop While ID > 0
   
   i = 14
   
   Куда_израсход_Код.SetFocus
   pathStr = ""
   Do
      If pathArray(i) = "" Then
      Else
         If pathStr = "" Then
            pathStr = pathArray(i)
         Else
            pathStr = pathStr & ", " & pathArray(i)
         End If
      End If
         i = i - 1
   Loop While i >= 0
   
   Куда_израсход_Код.Locked = False
   'сохраняем текстовое название предыдущего местонахождения прибора в переменную
   oldPathStr = Куда_израсход_Код.Text
   'записываем в поле Куда_израсход_Код текстовое название нового местонахождения
   Куда_израсход_Код.Text = pathStr
   rlStr = pathStr
   targetRelocChanged = True
   Куда_израсход_Код.Locked = True
   
   If clickCounter <> 0 Then
   Else           'случай совершения первого клика по дереву структуры
      Местонахождение.SetFocus
      Местонахождение.Locked = False
      'записываем в поле с меткой "Откуда перемещается" переменную с текстом предыдущего местонахождения
      Местонахождение.Text = oldPathStr
      Местонахождение.Locked = True
      'fromPlaceLabel.Caption = "Откуда перемещается"
      clickCounter = clickCounter + 1
   End If
   'фокусируемся на скрытом поле
   hidedTextBox.SetFocus
   'записываем его текущее значение - индекс, соответствующий текущему узлу структуры в переменную
   ID = Val(hidedTextBox.Text)
   'записываем в текстовое поле значение индекса структуры
   structure_idTextBox.SetFocus
   structure_idTextBox.Text = ID
   TrView.SetFocus
   
   'очистка
   rs.Close
   CurrentDb.Close
   
Exit Sub
ErrHandler:
   MsgBox Err.Description
End Sub 'TrView_Click


'обработка события "изменено содержание поля Куда израсходовано"
Private Sub Куда_израсход_Код_Dirty(Cancel As Integer)
   
   targetRelocChanged = True
   
End Sub 'Куда_израсход_Код_Dirty

'обнуление элементов массива, хранящего путь
Private Sub arrayInitilize(wpathArray() As String)
Dim i As Integer

   For i = 0 To 14
      wpathArray(i) = ""
   Next i
   
End Sub 'arrayInitilize

'процедура обновления в БД множества записей (по значению поля Количество)
Sub UpdateRecords()
On Error GoTo ErrorHandler
Dim dbs As DAO.Database
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset
Dim counter As Integer, recordsCounter As Long

   Set dbs = CurrentDb
   
   Set qdf = dbs.CreateQueryDef("UpdateRecordsQuery")
   qdf.SQL = Me.RecordSource
               
   Set rst = qdf.OpenRecordset()
   
   recordsCounter = 1
   rst.MoveFirst
   For counter = 1 To relocNumberTextBox.value
      If recordsCounter <> Me.CurrentRecord Then
         With rst
            .Edit
            '!Код_ТМЦ = auto
            !Откуда_перемещ = Местонахождение
            !Куда_израсход_Код = Куда_израсход_Код
            !Когда_израсходовано = relocationDateTextBox
            !Инициалы = personalTextBox
            !Причины = reasonComboBox
            !Состояние = stateComboBox
            !Количество = 1
            !notes = Примечание_Код
            !structure_id = Me.structure_idTextBox
            .Update
         End With
         rlNumber = rlNumber + 1
         'вставка строки в таблицу history_records, используя значения _
         из переменных хранения данных для исторической записи
         CurrentDb.Execute " INSERT INTO history_records(relocation_date,Куда_израсход_Код, " _
                            & "Инициалы,state_id,class_id,session_id,reason_id,Код_ТМЦ ) VALUES " _
                            & "('" & rDate & "','" _
                            & bRplace_id & "','" _
                            & brPersonal & "','" _
                            & bRstate_id & "','" _
                            & bCclass_id & "','" _
                            & bRsession_id & "','" _
                            & bRreason_id & "','" _
                            & rst.Fields(0) & " ' );"
      Else
      End If
      rst.MoveNext
      recordsCounter = recordsCounter + 1
   Next counter

    ' увеличение переменной для хранения количества перемещенных ТМЦ на единицу (из-за особенностей
    ' данного проекта одно ТМЦ всегда перемещается самой формой перемещения, так что если требуется
    ' переместить N ТМЦ, то N-1 перемещается чисто программно при помощи кода VB, а 1 шт всегда пере-
    ' мещается формой, так как форма основана на запросе Update, и при обновлении данных формы, они
    ' обновляются в таблицах, на которых основан запрос Update)
    rlNumber = rlNumber + 1
   'Return database to original.
   dbs.QueryDefs.Delete "UpdateRecordsQuery"

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

