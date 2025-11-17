VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ViewСклад"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim tr As Object

'добавление нового узла
Private Sub addNodeButton_Click()

End Sub 'addNodeButton_Click
'добавление нового ТМЦ
Private Sub addRecordButton_Click()

      DoCmd.OpenForm "InsertСклад", , , , , acDialog
      Me.subСкладUnnested.Requery

End Sub 'addRecordButton_Click

'нажатие кнопки "Запрос приборов во всех узлах"
Private Sub allQueryButton_Click()
Dim rst As Recordset

   Me.RecordSource = "SELECT * FROM cpyСклад "
   statusLabel.Caption = "Вывод ТМЦ всего дерева филиалов и организаций"
   '"популяризация" источника данных формы
   Set rst = Me.Recordset
   rst.MoveLast
   rst.MoveFirst

End Sub 'allQueryButton_Click
'очистка бланка материального пропуска
Private Sub clearPassBlankButton_Click()
On Error GoTo ErrHandler

   CurrentDb.Execute "DELETE * FROM passblank"
   MsgBox "Бланк очищен!", vbInformation, strMsgBox
   
Exit Sub
ErrHandler:
   MsgBox "Ошибка:" & Err.Description, vbCritical, "Очистка бланка"
End Sub 'clearPassBlankButton_Click

'нажатие копирования в макет матер. пропуска
Private Sub copyToPassBlankButton_Click()
On Error GoTo ErrHandler
Dim rst As Recordset, serial As Byte, i As Byte, recordsCount As Byte
     '  CurrentDb.Execute "CREATE TABLE passblank " _
     '             & "(Наименование TEXT, " _
     '             & "Тип TEXT, " _
     '             & "[Зав номер] TEXT, " _
     '             & "Количество LONG, " _
     '             & "Местонахождение TEXT); "
   Set rst = CurrentDb.OpenRecordset("SELECT Наименование FROM passblank")
   If rst.EOF And rst.BOF Then
      recordsCount = 0
   Else
      'populate the recordset
      rst.MoveLast
      rst.MoveFirst
      recordsCount = rst.RecordCount
   End If
   If recordsCount < 3 Then
      serial = recordsCount + 1
      CurrentDb.Execute "INSERT INTO passblank ( Наименование,Тип,[Зав номер],Количество,Местонахождение,Номер ) VALUES" _
                     & "('" & Наименование & _
                      "','" & Тип & _
                      "','" & Зав_номер & _
                      "','" & Количество & _
                      "','" & Местонахождение & _
                      "','" & serial & "');"
      MsgBox "ТМЦ скопировано!", vbInformation, strMsgBox
   Else
      MsgBox "В бланке уже 3 ТМЦ. Распечатайте бланк," _
         & "очистите его и заполните новый.", vbExclamation, strMsgBox
   End If
   
   rst.Close
Exit Sub
ErrHandler:
   MsgBox "Ошибка:" & Err.Description, vbCritical, strMsgBox
   rst.Close
End Sub 'copyToPassBlankButton_Click

' нажатие "Удалить текущее ТМЦ" (из буфера)
Private Sub deleteFromBufferButton_Click()
On Error GoTo ErrHandler
'Static SubRst As Recordset, counter As Long
Dim rst As DAO.Recordset
Dim strSearchName As String, Response As Integer

    Set rst = subСкладUnnested.Form.Recordset
    strSearchName = CStr(subСкладUnnested!Код_ТМЦ)
    rst.FindFirst "Код_ТМЦ = " & strSearchName
    If rst.NoMatch Then
        MsgBox "Запись не найдена", vbInformation + vbOKOnly, strMsgBox
    Else
        Response = MsgBox("Выбрано для удаления: " & vbCrLf _
                        & "Код ТМЦ: " & vbTab & rst.Fields(10) & vbCrLf _
                        & "Дата регистрации в БД: " & vbTab & rst.Fields(0) & vbCrLf _
                        & "Наименование: " & vbTab & rst.Fields(1) & vbCrLf _
                        & "Тип: " & vbTab & rst.Fields(2) & vbCrLf _
                        & "Инв. номер: " & vbTab & rst.Fields(3) & vbCrLf _
                        & "Зав. номер: " & vbTab & rst.Fields(4) & vbCrLf _
                        & "Количество: " & vbTab & rst.Fields(5) & vbCrLf _
                        & "Откуда: " & vbTab & rst.Fields(6) & vbCrLf _
                        & "Куда: " & vbTab & rst.Fields(7) & vbCrLf _
                        & "Основание: " & vbTab & rst.Fields(8) & vbCrLf _
                        & "Состояние: " & vbTab & rst.Fields(9) & vbCrLf & vbCrLf _
                        & "Удалить окончательно?", vbExclamation + vbOKCancel, strMsgBox)
        If Response = vbOK Then
            ' удаление ТМЦ из базы
            CurrentDb.Execute "DELETE * FROM Тип_Наличие WHERE " _
                             & "Тип_Наличие.Код_ТМЦ =" & strSearchName & ";"
            ' удаление истории движения ТМЦ
            CurrentDb.Execute "DELETE * FROM history_records WHERE " _
                             & "history_records.Код_ТМЦ =" & strSearchName & ";"
            MsgBox "Удаление произведено.", vbInformation + vbOKOnly, strMsgBox
        End If
        
    End If

   'обновить подчиненную форму буфера приборов после операций с базой данных
   Me.subСкладUnnested.Requery
   rst.MoveFirst

Exit Sub
ErrHandler:
   MsgBox Err.Description, , strMsgBox
End Sub 'deleteFromBufferButton_Click

'удаление узла из дерева
Private Sub deleteNodeButton_Click()

End Sub 'deleteNodeButton_Click
'нажатие кнопки Редактирование карточки ТМЦ
Private Sub editRecordButton_Click()

   DoCmd.OpenForm "EditForm", , , , , acDialog
   Me.Requery
   
End Sub 'editRecordButton_Click()

'закрытие формы - выход из приложения
Private Sub Form_Close()
Dim Noth As Variant

   'отметка окончания сессии идентифицированного пользователя
   Noth = session.setSessionEnd(auth.getSubdiv, auth.getLogin)
   DoCmd.Close acForm, "Authorization"
   'Application.Quit acQuitSaveAll

End Sub 'Form_Close


'загрузка формы
Private Sub Form_Load()
Dim rst As Recordset
Dim meRst As Recordset
Dim Noth As Variant
      
   'конструктор нового экземпляра класса clsTreeClass заполнения и работы с деревом
   Set tr = New clsTreeClass
   'инициализация контрола ТриВью на форме
   Set tr.Tree = Me.TrView.Object
   'использование методов класса clsTreeClass для инициализации переменных описывающих таблицу структуры дерева
   tr.Tbl = "Структура"
   tr.fldKey = "Код"
   tr.fldParent = "parent_id"
   tr.fldText = "Наименование"
   'заполнение дерева
   tr.GenerateTree
   
   'инициализация переменной Recordset и присвоение ей набора записей подчиненной формы (буфера ТМЦ)
   Set rst = subСкладUnnested.Form.Recordset
   'проверка условия, что в подчиненной форме (буфере ТМЦ) имеются записи
   If rst.BOF = False And rst.EOF = False Then
      relocateUnnestedButton.Enabled = True 'активировать кнопку перемещения ТМЦ из буфера
      'populate the recordset
      rst.MoveLast
      rst.MoveFirst
   End If
   
   
   'инициализация нового экземпляра класса, описывающего виды перемещений ТМЦ
   Set rl = New clsRelocateClass
   
   'инициализация метки статуса на форме просмотра ТМЦ
   statusLabel.Caption = "Вывод ТМЦ всего дерева филиалов и организаций"
   
   'инициализация нового экземпляра класса, идентифицирующего сессии работы с БД
   Set session = New clsSessionsClass
   
   ' инициализация нового экз. класса clsSelectRecordClass, описывающего
   ' методы и свойства процедур редактирования аттрибутов и выбора ТМЦ
   Set sr = New clsSelectRecordClass
   
   '"популяризация" источника данных формы
   Set meRst = Me.Recordset
   meRst.MoveLast
   meRst.MoveFirst
   
   'отметка начала запущенной сессии и идентификации пользователя
   Noth = session.setSessionBegin(auth.getSubdiv, auth.getLogin)
            
End Sub 'Form_Load
' нажатие кнопки "свойства принтера"
Private Sub getPaperListButton_Click()

   ' вывести все доступные принтеры в сообщении
   Call ShowPrinters
   ' вывести принтер по умолчанию и типы бумаги, которые он поддерживает
   Call GetPaperList

End Sub ' getPaperListButton_Click

'обработка нажатия кнопки на добавление всех путей из дерева структуры в таблицы "Куда_израсход" и
' "Откуда_перемещ"
Private Sub insertPathsButton_Click()
On Error GoTo ErrHandler
Dim rsMaster As Recordset, rsSlave As Recordset
Dim i As Integer, ID As Long, pathStr As String
Dim pathArray(14) As String
 
   Set rsMaster = CurrentDb.OpenRecordset("SELECT Код FROM Структура")
   
   Do Until rsMaster.EOF
      Call arrayInitilize(pathArray)
      i = 0
      ID = rsMaster("Код")
                 
      Do
         Set rsSlave = CurrentDb.OpenRecordset("SELECT Перевод,parent_id FROM Структура " _
              & " WHERE Код = " _
              & ID & "")
         ID = rsSlave("parent_id")
         pathArray(i) = rsSlave("Перевод")
         i = i + 1
      Loop While ID > 0
      rsSlave.Close
      
      i = 14
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
      CurrentDb.Execute " INSERT INTO Куда_израсход ( F27 ) VALUES " _
                      & " ('" & pathStr & "');"
      CurrentDb.Execute " INSERT INTO Откуда_перемещ ( F27 ) VALUES " _
                      & " ('" & pathStr & "');"
      rsMaster.MoveNext
   Loop

   'очистка
   If rsMaster.EOF Then
        rsMaster.Close
        MsgBox "All nodes exported successfully", vbInformation + vbOKOnly, strMsgBox
   Else
        rsSlave.Close
   End If
      
Exit Sub
ErrHandler:
   MsgBox Err.Description
End Sub 'insertPathsButton_Click

'обнуление элементов массива, хранящего путь
Private Sub arrayInitilize(wpathArray() As String)
Dim i As Integer

   For i = 0 To 14
      wpathArray(i) = ""
   Next i
   
End Sub 'arrayInitilize
'обработка нажатия кнопки запроса приборов в узле
Private Sub nestedQueryButton_Click()
On Error GoTo ErrHandler
Dim strNewRecord As String
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset
Dim a As DAO.QueryDef
   
   'обращаемся к скрытому полю формы, хранящему индекс текущего(со свойством Selected = True) узла
   hidedTextBox.SetFocus
   'проверка условия, что выбран какой-либо узел в дереве
   If hidedTextBox.Text <> "" Then
      'вызов ф-ии составления части строки запроса данных для выбранного узла
      strNewRecord = qwrParamsStr(Val(hidedTextBox.Text))
      'используем вычисленную часть строки в запросе с выборкой ТМЦ, лежащих в выбранном и дочерних узлах
      strNewRecord = "SELECT * FROM cpyСклад " _
           & " WHERE " _
           & strNewRecord & ""
      ' проверяем условие, что объект SomeRecordsQuery уже существует
      ' Такое возможно, если предыдущий запрос вернул не результат, а ошибку
      ' (например, если предыдущий запрос был слишком сложным)
      For Each a In CurrentDb.QueryDefs
        If a.Name = "SomeRecordsQuery" Then
            Debug.Print a.Name
            CurrentDb.QueryDefs.Delete "SomeRecordsQuery"
        End If
        Debug.Print a.Name
      Next
            
      Set qdf = CurrentDb.CreateQueryDef("SomeRecordsQuery")
      qdf.SQL = strNewRecord
      'открываем набор данных, имеющих отношение к выбранному и дочерним узлам
      Set rst = qdf.OpenRecordset()
      If rst.BOF Or rst.EOF Then
         Me.RecordSource = "nullQuery"   'если запрос не вернул данных, выводим пустые поля
      Else
         Me.RecordSource = strNewRecord 'либо - выводим результат запроса
      End If
   
      'Очистка базы данных
      CurrentDb.QueryDefs.Delete "SomeRecordsQuery"
      rst.Close
      qdf.Close
      CurrentDb.Close
      
      'фокусируемся на дереве
      TrView.SetFocus
   Else
      MsgBox "Не выбран ни один узел дерева структуры!", vbExclamation, strMsgBox
      Exit Sub
   End If
     
   'инициализация метки статуса на форме просмотра ТМЦ
   statusLabel.Caption = "Вывод ТМЦ выделенного узла дерева и вложенных в нем узлов"
   
Exit Sub
'обработчик ошибок
ErrHandler:
   MsgBox Err.Description
End Sub 'nestedQueryButton_Click

'обработка нажатия кнопки справочника наименований
Private Sub OpenNamesGlossaryButton_Click()

   DoCmd.OpenForm FormName:="NamesGlossary", datamode:=acFormReadOnly, WindowMode:=acDialog
   
End Sub 'OpenNamesGlossaryButton_Click
'обработка нажатия кнопки справочника типов
Private Sub openTypesGlossarybutton_Click()

   DoCmd.OpenForm FormName:="TypesGlossary", datamode:=acFormReadOnly, WindowMode:=acDialog

End Sub 'openTypesGlossarybutton_Click
'нажатие кнопки Печать отчета ЗИП АСУ
Private Sub printAsuReportButton_Click()
   
   If enabPrinting Then
      DoCmd.OpenReport "AsuReport", acViewNormal
   End If

End Sub 'printAsuReportButton_Click
'нажатие "Печатать матер. пропуск"
Private Sub printPassBlankButton_Click()

   If enabPrinting Then
      DoCmd.OpenReport "passblankReport"
   End If

End Sub 'printPassBlankButton_Click

'переместить текущее ТМЦ в пределах структуры
Private Sub relocateButton_Click()

   'метод Inside - перемещение структура - структура
   rl.Inside
   rl.inRelocation = True
   DoCmd.OpenForm "UpdateСклад", , , , , acDialog
   
   Me.Requery
   If Me.Recordset.RecordCount = 0 Then
         Me.RecordSource = "nullQuery"
   End If
 
End Sub 'relocateButton_Click
'нажатие кнопки перемещения ТМЦ из структуры в буфер приборов
Private Sub relocateToBufButton_Click()

   'метод InsideOute - перемещение структура - буфер
   rl.InsideOut
   rl.insOutRelocation = True
   DoCmd.OpenForm "UpdateСклад", , , , , acDialog
   
   Me.Requery
   If Me.Recordset.RecordCount = 0 Then
      Me.RecordSource = "nullQuery"
   End If
   Me.subСкладUnnested.Requery
   
End Sub 'relocateToBufButton_Click

'нажатие кнопки перемещения ТМЦ из буфера приборов в структуру
Private Sub relocateUnnestedButton_Click()
On Error GoTo ErrHandler
Static SubRst As Recordset, counter As Long
   
   Set SubRst = subСкладUnnested.Form.Recordset
   'метод Outside - перемещение буфер - структура (для новых или перемещаемых между филиалами ТМЦ)
   queryString = rl.Outside(SubRst)
   rl.outRelocation = True
   DoCmd.OpenForm "UpdateСклад", , , , , acDialog

   'обновить подчиненную форму буфера приборов после перемещения оттуда
   Me.subСкладUnnested.Requery
   SubRst.MoveFirst

Exit Sub
ErrHandler:
   MsgBox Err.Description

End Sub 'relocateUnnestedButton_Click

'кнопка вывода отчета по складу в Ексель
Private Sub reportSkladToExcelbutton_Click()
On Error GoTo ErrorHandler

   'подготовить отчет на основе шаблона из запроса Warehouse
   'создать таблицу-копию отчета по складу
   Call CreateTable(PrepareWarehouseReport)
   'экспорт Отчета в XML-файл
   Application.ExportXML acExportTable, "Отчет по Складу", _
      "Отчет по Складу.xml", "Отчет по Складу.xsd"
      
   'удалить таблицу-копию отчета по складу
   CurrentDb.Execute "DROP TABLE [Отчет по Складу];"

   'вызов процедуры работы с XML-файлом отчета
   Call WorkWithXML
   
   Kill "Отчет по Складу.xml"
   Kill "Отчет по Складу.xsd"
   
   MsgBox "Перечень складских ТМЦ выведен в файл: Мои документы\" & _
      "Отчет по Складу на " & CStr(Date) & ".xlsx", vbInformation + vbOKOnly, strMsgBox

   'DoCmd.OutputTo acOutputQuery, "Отчет по складу", acFormatXLSX, _
   '   "Отчет по Складу на " & CStr(Date) & ".xlsx"
   'MsgBox "Перечень складских ТМЦ выведен в файл: Мои документы\" & _
   '   "Отчет по Складу на " & CStr(Date) & ".xlsx"
Exit Sub

ErrorHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   'If Err.number <> 1004 Then
   '   MsgBox "Произошла ошибка доступа к файлу!"
   'End If
End Sub 'reportSkladToExcelbutton_Click
'нажатие кнопки Отчет АСУ в PDF
Private Sub saveAsuReportAsPdfButton_Click()
On Error GoTo ErrorHandler

   DoCmd.OutputTo acOutputReport, "AsuReport", _
      acFormatRTF, "ЗИП и склад ТМЦ АСУ на " & CStr(Date) & ".rtf", True
   MsgBox "Бланк выведен в 'Мои документы/ЗИП и склад ТМЦ АСУ на" & CStr(Date) & ".rtf'", vbInformation + vbOKOnly, strMsgBox
   
Exit Sub

ErrorHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   If Err.number = 31036 Then
      MsgBox "Произошла ошибка вывода в файл. Возможно не устанолен компонент экспорта в Adobe", vbCritical + vbOKOnly, strMsgBox
   End If
   Exit Sub

End Sub 'saveAsuReportAsPdfButton_Click

'клик на узле дерева структуры предприятий
Private Sub TrView_Click()
   
   'запись индекса выбранного узла в скрытом поле на форме
   hidedTextBox.value = tr.GetKey
   
End Sub 'TrView_Click

'Создание объекта Excel и открытие в нем экспортированного xml файла "Отчет по складу"
Sub WorkWithXML()
On Error GoTo ErrorHandler

Dim objExcel As Excel.Application
Dim objWS As Excel.Worksheet

   Set objExcel = New Excel.Application
   'objExcel.Workbooks.Add
   'Set objWS = objExcel.ActiveSheet
   
   objExcel.Workbooks.OpenXML _
      fileName:="Отчет по Складу.xml", _
      LoadOption:=xlXmlLoadImportToList
   objExcel.ActiveSheet.Name = "Отчет по складу"
   objExcel.ActiveWorkbook.SaveAs "Отчет по Складу на " & CStr(Date) & ".xlsx"
        
   objExcel.Quit
   
Exit Sub

ErrorHandler:
   If Err.number <> 1004 Then
      MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
      MsgBox "Произошла ошибка сохранения или вывода в файл! ", vbCritical + vbOKOnly, strMsgBox
   End If
Exit Sub
  
End Sub 'WorkWithXML

'процедура создания таблицы-копии отчета по складу "Warehouse"
Private Sub CreateTable(whrCls As String)

    CurrentDb.Execute "CREATE TABLE [Отчет по складу] " _
               & "([Дата получения] DATETIME, Наименование TEXT, " _
               & "[Тип, марка] TEXT, " _
               & "[Инв номер] TEXT, " _
               & "[Зав номер] TEXT, " _
               & "[Диапазон, типоразмер] TEXT, " _
               & "[Состояние] TEXT, " _
               & "[Кол-во на отчетную дату] LONG, " _
               & "Местонахождение TEXT, " _
               & "Примечание TEXT );"

   CurrentDb.Execute "INSERT INTO [Отчет по складу] " _
               & "SELECT Warehouse.[Дата получения],Warehouse.Наименование," _
               & "Warehouse.[Тип, марка],Warehouse.[Инв номер],Warehouse.[Зав номер], " _
               & "Warehouse.[Диапазон, типоразмер],Warehouse.[Кол-во на отчетную дату]," _
               & "Warehouse.Местонахождение,Warehouse.Состояние,Warehouse.Примечание  " _
               & "FROM Warehouse  " _
               & "WHERE " & whrCls & ";"
        
End Sub 'CreateTable


'обновить содержимое буфера приборов
Private Sub unnestedQueryButton_Click()
   
   Me.subСкладUnnested.Requery
   
End Sub 'unnestedQueryButton_Click(

'нажатие кнопки вывода отчета по использованию - в Ексель
Private Sub useReportButton_Click()
'вывод отчета по использованию за последние 31 календарный день

   'вывести запрос "UseReport" в Excel
   DoCmd.OutputTo acOutputQuery, "UseReport", acFormatXLSX, _
      "Отчет по использованию ТМЦ " & ".xlsx" 'за период с " & CStr(Date - 31) _
      & " по " & CStr(Date) & ".xlsx"
   MsgBox "Отчет по использованию выведен в файл: Мои документы\" & _
   "Отчет по использованию ТМЦ " & ".xlsx", vbInformation + vbOKOnly, strMsgBox
   'за период с " & CStr(Date - 31) _
      & " по " & CStr(Date) & ".xlsx" ', vbInformation + vbOKOnly, strMsgBox

End Sub 'useReportButton_Click
'функция подготовки отчета по складу из шаблона запроса Warehouse
Private Function PrepareWarehouseReport() As String
Dim idsRst As Recordset, whrClause As String
Dim enterCounter As Integer

   'счетчик вхождений в ветвь "Да" блока if/else
   enterCounter = 0
   'вычисляемая строка
   whrClause = "structure_id = "
   'запрос всех кодов узлов структуры, отмеченных как места хранения
   Set idsRst = CurrentDb.OpenRecordset("SELECT Код FROM Структура WHERE store = TRUE")
   idsRst.MoveLast
   idsRst.MoveFirst
   Do Until idsRst.EOF
      If IsItMyChild(idsRst.Fields(0)) Then
         enterCounter = enterCounter + 1
         'если не первое вхождение в эту ветвь блока if/else
         If enterCounter <> 1 Then
            whrClause = whrClause & " OR " & "structure_id = " & CStr(idsRst.Fields(0))
         Else
            whrClause = whrClause & CStr(idsRst.Fields(0))
         End If
      Else
      End If
      idsRst.MoveNext
   Loop
   
   'если в текущем филиале нет мест хранения _
   'вывести пустой отчет
   If enterCounter = 0 Then
      whrClause = whrClause & "1"
   End If
   
   'возврат вычисленной строки
   PrepareWarehouseReport = whrClause
    
End Function 'PrepareWarehouseReport

'процедура определения, является ли переданный параметром код, _
'кодом места хранения, принадлежащего активному филиалу
Private Function IsItMyChild(ByVal ID As Long) As Boolean

   'если переданный id является дочерним узлом для текущего филиала
   If auth.getId = getDesiredSubdivCode(ID) Then
      IsItMyChild = True
   Else
      IsItMyChild = False
   End If
   
End Function 'IsItMyChild2
'нажатие кнопки "просмотреть бланк матер. пропуска"
Private Sub viewPassBlankButton_Click()

   DoCmd.OpenReport "passblankReport", acViewPreview, , , acDialog

End Sub 'viewPassBlankButton_Click

' нажатие кнопки "Отчет по сессиям работы"
Private Sub viewSessionsButton_Click()
On Error GoTo ErrorHandler
Dim Response As Integer

    DoCmd.OpenReport "sessionsReport", acViewReport, , , acDialog
    Response = MsgBox("Следует ли вывести этот отчет в файл Rich Text Format?", vbQuestion + vbYesNo, strMsgBox)
    If Response = vbYes Then
        DoCmd.OutputTo acOutputReport, "sessionsReport", _
                       acFormatRTF, "Отчет по сессиям на " & CStr(Date) & ".rtf", True
        MsgBox "Отчет выведен в 'Мои документы/Отчет по сессиям на " & CStr(Date) & ".rtf", vbInformation + vbOKOnly, strMsgBox
    End If
   
Exit Sub

ErrorHandler:
   MsgBox "Ошибка #: " & Err.number & vbCrLf & vbCrLf & Err.Description
   If Err.number = 31036 Then
      MsgBox "Произошла ошибка вывода в файл. Возможно не устанолен компонент экспорта в Adobe", vbCritical + vbOKOnly, strMsgBox
   End If
   Exit Sub

End Sub 'viewSessionsButton_Click


