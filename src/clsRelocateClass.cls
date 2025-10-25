VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelocateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'класс, описывающий объект "Перемещение"                                                        '
'класс моделирует оператора, определяющего свойства и метода выполняемой операции перемещения   '
'ТМЦ, а также контролирующего ход этой операции                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Dim rlc As clsRelocateClass
'объявление переменных для хранения свойств
Dim inProperty As Boolean, outProperty As Boolean, insOutProperty As Boolean
Dim toPositionProperty As Boolean, toStoreProperty As Boolean, toUsed_StoreProperty As Boolean
Dim busyPositionProperty As Boolean
Dim rlcToStoreProperty As Boolean, rlcToPositionProperty As Boolean, rlcToUsed_StoreProperty As Boolean
Dim recsCountProperty As Long

'метод, который при типе перемещения "структура - структура" вычисляет строчную переменную,
'используемую в запросе. Этот запрос является затем источником данных формы UpdateСклад
Public Sub Inside()

   queryString = defineInsideQuery
   
End Sub 'Inside()
'метод, который при типе перемещения структура - буфер вычисляет строчную переменную,
'используемую в запросе.
Public Sub InsideOut()

   queryString = defineInsideQuery
   
End Sub 'InsideOut
'метод, который при типе перемещения буфер - структура вычисляет строчную переменную,
'используемую в запросе.
Public Function Outside(wrst As Recordset) As String

   Outside = defineOutsideQuery(wrst)
   
End Function 'Outside

'метод, проверяющий вид перемещения "на позицию". Также устанавливает сам соответствующее
'свойство toPosition объекта Перемещение
Public Function isItPosition(ID As Long) As Boolean
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT position FROM Структура WHERE Код = " & ID & "")
   isItPosition = rst.Fields(0)
   toPosition = rst.Fields(0)
   If toPosition Then
      Call isPositionBusy(ID)
   End If
   rst.Close
      
End Function 'isItPosition
'метод, проверяющий вид перемещения "на место хранения". Также устанавливает сам соответствующее
'свойство toStore объекта Перемещение
Public Function isItStore(ID As Long) As Boolean
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT store FROM Структура WHERE Код = " & ID & "")
   isItStore = rst.Fields(0)
   toStore = rst.Fields(0)
   rst.Close
      
End Function 'isItStore
'метод, проверяющий вид перемещения "на место хранения". Также устанавливает сам соответствующее
'свойство toStore объекта Перемещение
Public Function isItUsed_Store(ID As Long) As Boolean
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT used_store FROM Структура WHERE Код = " & ID & "")
   isItUsed_Store = rst.Fields(0)
   toUsed_Store = rst.Fields(0)
   rst.Close
      
End Function 'isItUsed_Store

'метод, проверяющий занятость занимаемой позиции
Public Sub isPositionBusy(wId As Long)
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT Код_ТМЦ FROM UpdateСклад WHERE structure_id = " & wId & "")
   If rst.RecordCount = 1 Then
      busyPosition = True
   End If

End Sub 'isPositionBusy
'установка значения свойства, которое является признаком перемещения "структура - структура"
Property Let inRelocation(value As Boolean)

   If value Then
      inProperty = True
   Else
      inProperty = False
   End If
   
End Property 'inRelocation
'возврат значения свойства, которое является признаком перемещения "структура - структура"
Property Get inRelocation() As Boolean

   If inProperty Then
      inRelocation = True
   Else
      inRelocation = False
   End If
   
End Property 'inRelocation
'установка значения свойства, которое является признаком перемещения "буфер - структура"
Property Let outRelocation(value As Boolean)

   If value Then
      outProperty = True
   Else
      outProperty = False
   End If
   
End Property 'inRelocation
'возврат значения свойства, которое является признаком перемещения "буфер - структура"
Property Get outRelocation() As Boolean

   If outProperty Then
      outRelocation = True
   Else
      outRelocation = False
   End If
   
End Property 'inRelocation
'установка значения свойства, которое является признаком перемещения "структура - буфер"
Property Let insOutRelocation(value As Boolean)

   If value Then
      insOutProperty = True
   Else
      insOutProperty = False
   End If
   
End Property 'inRelocation
'возврат значения свойства, которое является признаком перемещения "структура - буфер"
Property Get insOutRelocation() As Boolean

   If insOutProperty Then
      insOutRelocation = True
   Else
      insOutRelocation = False
   End If
   
End Property 'inRelocation
'установка значения свойства, которое является признаком перемещения "на позицию"
Property Let toPosition(value As Boolean)

   If value Then
      toPositionProperty = True
   Else
      toPositionProperty = False
   End If
   
End Property 'toPosition
'возврат значения свойства, которое является признаком перемещения "на позицию"
Property Get toPosition() As Boolean

   If toPositionProperty Then
      toPosition = True
   Else
      toPosition = False
   End If
   
End Property 'toPosition
'установка значения свойства, которое является признаком перемещения "на место хранения"
Property Let toStore(value As Boolean)

   If value Then
      toStoreProperty = True
   Else
      toStoreProperty = False
   End If
   
End Property 'toStore
'возврат значения свойства, которое является признаком перемещения "на место хранения"
Property Get toStore() As Boolean

   If toStoreProperty Then
      toStore = True
   Else
      toStore = False
   End If
   
End Property 'toStore
'установка значения свойства, которое является признаком перемещения "в использованные"
Property Let toUsed_Store(value As Boolean)

   If value Then
      toUsed_StoreProperty = True
   Else
      toUsed_StoreProperty = False
   End If
   
End Property 'toUsed_Store
'возврат значения свойства, которое является признаком перемещения "в использованные"
Property Get toUsed_Store() As Boolean

   If toUsed_StoreProperty Then
      toUsed_Store = True
   Else
      toUsed_Store = False
   End If
   
End Property 'toUsed_Store

'установка значения свойства, которое является признаком, что позиция уже занята прибором
Property Let busyPosition(value As Boolean)

   If value Then
      busyPositionProperty = True
   Else
      busyPositionProperty = False
   End If
   
End Property 'busyPosition
'возврат значения свойства, которое является признаком, что позиция уже занята прибором
Property Get busyPosition() As Boolean

   If busyPositionProperty Then
      busyPosition = True
   Else
      busyPosition = False
   End If
   
End Property 'busyPosition
'установка значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на склад
Property Let rlcToStore(value As Boolean)

   If value Then
      rlcToStoreProperty = True
   Else
      rlcToStoreProperty = False
   End If
   
End Property 'rlcToStore
'возврат значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на склад
Property Get rlcToStore() As Boolean

   If rlcToStoreProperty Then
      rlcToStore = True
   Else
      rlcToStore = False
   End If
   
End Property 'rlcToStore
'установка значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на позицию
Property Let rlcToPosition(value As Boolean)

   If value Then
      rlcToPositionProperty = True
   Else
      rlcToPositionProperty = False
   End If
   
End Property 'rlcToPosition
'возврат значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на позицию
Property Get rlcToPosition() As Boolean

   If rlcToPositionProperty Then
      rlcToPosition = True
   Else
      rlcToPosition = False
   End If
   
End Property 'rlcToPosition
'установка значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на склад
Property Let rlcToUsed_Store(value As Boolean)

   If value Then
      rlcToUsed_StoreProperty = True
   Else
      rlcToUsed_StoreProperty = False
   End If
   
End Property 'rlcToUsed_Store
'возврат значения свойства, которое является признаком, что для
'ТмЦ разрешено перемещение на склад
Property Get rlcToUsed_Store() As Boolean

   If rlcToUsed_StoreProperty Then
      rlcToUsed_Store = True
   Else
      rlcToUsed_Store = False
   End If
   
End Property 'rlcToUsed_Store
'установка значения свойства, которое является счетчиком записей
Property Let recsCount(value As Long)

   recsCountProperty = value
   
End Property 'recsCount
'возврат значения свойства, которое является счетчиком записей
Property Get recsCount() As Long

   recsCount = recsCountProperty

End Property 'rlcToUsed_Store


'функция вычисления строки запроса при типе перемещения "структура-структура", "структура-буфер"
Private Function defineInsideQuery() As String
Dim namePribor As String, typePribor As String, serNumPribor As String, ID As Long
Dim Query As String

      'создание переменной Наименования прибора для запроса
      namePribor = "*" & Forms("ViewСклад").Наименование.value & "*"
      'создание переменной Типа прибора для запроса
      Forms("ViewСклад").Тип.SetFocus
      If Forms("ViewСклад").Тип.Text <> "" Then
         typePribor = "*" & Forms("ViewСклад").Тип.value & "*"
      Else
         typePribor = ""
      End If
      'создание переменной Зав.номера прибора для запроса
      Forms("ViewСклад").Зав_номер.SetFocus
      If Forms("ViewСклад").Зав_номер.Text <> "" Then
         serNumPribor = "*" & Forms("ViewСклад").Зав_номер.value & "*"
      Else
         serNumPribor = ""
      End If
      'запись идентификатора положения прибора в структуре в переменную
      ID = Forms("ViewСклад").structure_idTextBox.value
      'условие что и тип, и зав. номер - не пустые
      If typePribor <> "" And serNumPribor <> "" Then
         Query = "SELECT * FROM UpdateСклад " _
                 & " WHERE Наименование Like '" _
                 & namePribor & "' " _
                 & " AND Тип Like '" _
                 & typePribor & "' " _
                 & " AND Зав_номер Like '" _
                 & serNumPribor & "' " _
                 & " AND structure_id = " _
                 & ID & ""
      Else
         'условие, что тип - пустой, а зав. номер - не пустой
         If typePribor = "" And serNumPribor <> "" Then
            typePribor = "Is Null"
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип " _
                    & typePribor & "" _
                    & " AND Зав_номер Like '" _
                    & serNumPribor & "' " _
                    & " AND structure_id = " _
                    & ID & ""
         End If
         'условие, что тип - не пустой, а зав.номер - пустой
         If typePribor <> "" And serNumPribor = "" Then
            serNumPribor = """"""
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип Like '" _
                    & typePribor & "' " _
                    & " AND Зав_номер = " _
                    & serNumPribor & "" _
                    & " AND structure_id = " _
                    & ID & ""
         End If
         'условие, что и тип, и зав.номер- пустые
         If typePribor = "" And serNumPribor = "" Then
            serNumPribor = """"""
            typePribor = "Is Null"
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип " _
                    & typePribor & " " _
                    & " AND Зав_номер = " _
                    & serNumPribor & " " _
                    & " AND structure_id = " _
                    & ID & ""
         End If
      End If
      
      'источником данных формы сделать запрос приборов, соответствующи типу, наименованию,
      'зав. номеру текущей записи формы ViewСклад

      defineInsideQuery = Query
      
End Function 'defineInsideQuery
'функция вычисления строки запроса при типе перемещения "буфер-структура"
Private Function defineOutsideQuery(wwrst As Recordset) As String
Dim namePribor As String, typePribor As String, serNumPribor As String, ID As String
Dim Query As String
   
      'wwrst.MoveNext
      'wwrst.MovePrevious
      'создание переменной Наименования прибора для запроса
      namePribor = "*" & wwrst.Fields(1) & "*"
      'создание переменной Типа прибора для запроса
      If wwrst.Fields(2) <> "" Then
         typePribor = "*" & wwrst.Fields(2) & "*"
      Else
         typePribor = ""
      End If
      'создание переменной Зав.номера прибора для запроса
      If wwrst.Fields(4) <> "" Then
         serNumPribor = "*" & wwrst.Fields(4) & "*"
      Else
         serNumPribor = ""
      End If
      'идентификатор положения прибора, в буфере всегда равен NULL
      ID = "Is Null"
      'условие что и тип, и зав. номер - не пустые
      If typePribor <> "" And serNumPribor <> "" Then
         Query = "SELECT * FROM UpdateСклад " _
                 & " WHERE Наименование Like '" _
                 & namePribor & "' " _
                 & " AND Тип Like '" _
                 & typePribor & "' " _
                 & " AND Зав_номер Like '" _
                 & serNumPribor & "' " _
                 & " AND structure_id " _
                 & ID & ""
      Else
         'условие, что тип - пустой, а зав. номер - не пустой
         If typePribor = "" And serNumPribor <> "" Then
            typePribor = "Is Null"
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип " _
                    & typePribor & "" _
                    & " AND Зав_номер Like '" _
                    & serNumPribor & "' " _
                    & " AND structure_id " _
                    & ID & ""
         End If
         'условие, что тип - не пустой, а зав.номер - пустой
         If typePribor <> "" And serNumPribor = "" Then
            serNumPribor = """"""
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип Like '" _
                    & typePribor & "' " _
                    & " AND Зав_номер = " _
                    & serNumPribor & "" _
                    & " AND structure_id " _
                    & ID & ""
         End If
         'условие, что и тип, и зав.номер- пустые
         If typePribor = "" And serNumPribor = "" Then
            serNumPribor = """"""
            typePribor = "Is Null"
            Query = "SELECT * FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип " _
                    & typePribor & " " _
                    & " AND Зав_номер = " _
                    & serNumPribor & " " _
                    & " AND structure_id " _
                    & ID & ""
         End If
      End If
      
      'источником данных формы сделать запрос приборов, соответствующи типу, наименованию,
      'зав. номеру текущей записи формы subСкладUnnested

      defineOutsideQuery = Query
      
End Function 'defineOutsideQuery




