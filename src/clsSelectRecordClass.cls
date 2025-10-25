VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelectRecordClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Класс, определяющий:
' 1. Методы вычисления строковых переменных, используемых при выборе определенных записей                                    '                       '
'    определенных источников данных.
' 2. Свойства, хранящие информацию о причине редактирования аттрибутов ТМЦ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim src As clsSelectRecordClass
' переменная для хранения текста причины редактирования аттрибутов
Dim reasonString As String
' переменная для хранения номера причины редактирования аттрибутов
Dim reason_id As Integer
' метод, вычисляющий строчную переменную, используемую
' для выбора только определенных записей из источника данных
Public Sub SelInsertСклад()

   ' использование оператора метода Inside,
   ' выполняющего требуемую работу
   queryString = SelInsertQuery
   
End Sub 'SelectRecord()

' функция вычисления строки запроса для выбора определенног ТМЦ из
' источника данных
Private Function SelInsertQuery() As String
Dim namePribor As String, typePribor As String, serNumPribor As String, ID As Long
Dim Query As String
Dim rst As Recordset, oneId As Boolean

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
         Query = "SELECT Код_ТМЦ FROM UpdateСклад " _
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
            Query = "SELECT Код_ТМЦ FROM UpdateСклад " _
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
            serNumPribor = "Is Null"
            Query = "SELECT Код_ТМЦ FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип Like '" _
                    & typePribor & "' " _
                    & " AND Зав_номер " _
                    & serNumPribor & "" _
                    & " AND structure_id = " _
                    & ID & ""
         End If
         'условие, что и тип, и зав.номер- пустые
         If typePribor = "" And serNumPribor = "" Then
            serNumPribor = "Is Null"
            typePribor = "Is Null"
            Query = "SELECT Код_ТМЦ FROM UpdateСклад " _
                    & " WHERE Наименование Like '" _
                    & namePribor & "' " _
                    & " AND Тип " _
                    & typePribor & " " _
                    & " AND Зав_номер " _
                    & serNumPribor & " " _
                    & " AND structure_id = " _
                    & ID & ""
         End If
      End If

   ' создать источник данных на основе вычисленной строки
   Set rst = CurrentDb.OpenRecordset(Query)
   rst.MoveFirst
   ' строка запроса модифицирована для использования с источником данных (запросом) InsertСклад
   Query = "SELECT * FROM InsertСклад WHERE "
   ' признак единственного ТМЦ в запросе
   oneId = True
   ' вычисление строки запроса в зависимости от количества строк источника данных rst
   Do
      If oneId Then
         Query = Query & "Код_ТМЦ = " & rst(0)
      Else
         Query = Query & " OR " & "Код_ТМЦ = " & rst(0)
      End If
      oneId = False
      rst.MoveNext
   Loop Until rst.EOF
   ' функция вернет значение вычисленной строки запроса
   SelInsertQuery = Query

End Function ' SelInsertQuery
' возврат значения свойства "Номер Причины"
Public Property Get ReasonId() As Variant

End Property ' ReasonId
' установка свойства "Номер Причины"
Public Property Let ReasonId(ByVal vNewValue As Variant)

   reason_id = vNewValue
        
End Property ' ReasonId
' возврат значения свойства "Текст причины"
Public Property Get ReasonText() As Variant

   ReasonText = reasonString
 
End Property ' ReasonText
' установка свойства "Текст причины"
Public Property Let ReasonText(ByVal vNewValue As Variant)

   reasonString = vNewValue
   
End Property ' ReasonText
