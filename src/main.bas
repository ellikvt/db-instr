Attribute VB_Name = "main"
Option Compare Database
Option Explicit

'экземпляр класса clsAuthorizationClass
Public auth As Object
'экземпляр класса clsSessionsClass
Public session As Object
'экземпляр класса clsRelocateClass
Public rl As Object
' экземпляр класса clsSelectRecordClass
Public sr As Object
Public cancelUpdateQwr As Boolean
'счетчик количества записей
Public recordsCount As Long
'переменная для записи запроса при перемещении ТМЦ
Public queryString As String
' сообщение программы (пишется в заглавиях окон MsgBox)
Public Const strMsgBox As String = "Информатор КИП"
' код таблицы Откуда_перемещ, соответствующий полю F27 со значением "введено в базу данных"
Public Const SOURCE_ID = 147
' код таблицы Куда_израсход, соответствующий полю F27 со значением "буфер приборов"
Public Const DESTINATION_ID = 226

'объявление вызова функции получения имени локального компьютера Win API
Declare Function GetComputerName _
 Lib "kernel32" Alias "GetComputerNameA" _
  (ByVal lpBuffer As String, nSize As Long) _
   As Long
   
' Declaration for the DeviceCapabilities function API call.
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
    ByVal lpDevMode As Long) As Long
    
' DeviceCapabilities function constants.
Private Const DC_PAPERNAMES = 16
Private Const DC_PAPERS = 2
Private Const DC_BINNAMES = 12
Private Const DC_BINS = 6
Private Const DEFAULT_VALUES = 0
   
   
   
'инициализация глобальных переменных
Public Sub initVarsInMain()

   cancelUpdateQwr = False
   recordsCount = 0
   queryString = ""
   
End Sub 'initVarsInMain

'получение имени локального компьютера с помощью вызова Windows API
Public Function GetThisComputerName() As String
Dim strName As String
Dim lngChars As Long
Dim lngRet As Long

   strName = Space(255)
   lngChars = 255
   
   lngRet = GetComputerName(strName, lngChars)
   If lngRet > 0 Then
      GetThisComputerName = Left(strName, lngChars)
   Else
      GetThisComputerName = "Невозможно получить имя."
   End If

End Function 'GetThisComputerName

'Рекурсивная функция
'Возвращает строку, продолжающую оператор WHERE запроса ТМЦ
'в дочерних узлах, для выбранного в тривью родительского узла
Public Function qwrParamsStr(value As Long) As String
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT Код,parent_id FROM Структура " _
                                    & "WHERE parent_id = " & value & "")
   If rst.RecordCount <> 0 Then
      'блок рекурсивного вызова
      rst.MoveFirst
      Do Until rst.EOF
         If qwrParamsStr <> "" Then
            qwrParamsStr = qwrParamsStr & " OR " & qwrParamsStr(rst.Fields(0)) & " OR structure_id = " & CStr(value)
         Else
            qwrParamsStr = qwrParamsStr(rst.Fields(0)) & " OR structure_id = " & CStr(value)
         End If
         rst.MoveNext
      Loop
   Else
      'конечный шаг рекурсии
      qwrParamsStr = "structure_id = " & CStr(value)
   End If
   
   rst.Close
   
End Function 'qwrParamsStr
'Возвращает Код из таблицы Структура, соответствующий тому филиалу, вложенный узел которого передан в качестве
'параметра value. Если в качестве параметра передан сам родитель (филиал), то возвращается его Код
Public Function getDesiredSubdivCode(value As Long) As Long
Dim rst As Recordset, Code As Long, condition As Boolean

   Set rst = CurrentDb.OpenRecordset("SELECT Код,parent_id,subdivision FROM Структура " _
                                    & "WHERE Код = " & value & "")
                                    
   Do Until rst.Fields(2) = True
      Set rst = CurrentDb.OpenRecordset("SELECT Код,parent_id,subdivision FROM Структура " _
                                    & "WHERE Код = " & rst.Fields(1) & "")
   Loop
   getDesiredSubdivCode = rst.Fields(0)
      
   rst.Close
   
End Function 'getDesiredSubdivCode
' функция, которая:
' a. Использует функцию Windows API DeviceCapabilities для вывода сообщения
' б. Формирует сообщение на основании п. а) о принтере по умолчанию
' в. Формирует сообщение на основании п. а) о списке поддерживаемых типов бумаги
Sub GetPaperList()
    Dim lngPaperCount As Long
    Dim lngCounter As Long
    Dim hPrinter As Long
    Dim strDeviceName As String
    Dim strDevicePort As String
    Dim strPaperNamesList As String
    Dim strPaperName As String
    Dim intLength As Integer
    Dim strMsg As String
    Dim aintNumPaper() As Integer
    
    On Error GoTo GetPaperList_Err
    
    ' Получаем имя и порт принтера по умолчанию.
    strDeviceName = Application.Printer.DeviceName
    strDevicePort = Application.Printer.Port
    
    ' Get the count of paper names supported by the printer.
    ' Получаем количество наименований бумаги, поддерживаемых принтером
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERNAMES, _
        lpOutput:=ByVal vbNullString, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Re-dimension the array to the count of paper names.
    ' Меняем размер массива для наименований бумаги
    ReDim aintNumPaper(1 To lngPaperCount)
    
    ' Pad the variable to accept 64 bytes for each paper name.
    ' Присваиваем переменную 64 бита для хранения каждого наименования бумаги
    strPaperNamesList = String(64 * lngPaperCount, 0)

    ' Get the string buffer of all paper names supported by the printer.
    ' Получаем строковый буфер для всех наименований бумаги, поддерживаемых принтером
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERNAMES, _
        lpOutput:=ByVal strPaperNamesList, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Get the array of all paper numbers supported by the printer.
    ' Получаем массив всех номеров наименований, поддерживаемых принтером.
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERS, _
        lpOutput:=aintNumPaper(1), _
        lpDevMode:=DEFAULT_VALUES)
    
    ' List the available paper names.
    ' листинг доступных наименований бумаги.
    strMsg = "В настоящий момент используется принтер: " & strDeviceName & _
      "." & vbCrLf & "Для него доступны следующие типы бумаги:" & vbCrLf
    For lngCounter = 1 To lngPaperCount
        
        ' Parse a paper name from the string buffer.
        ' Вставить наименование бумаги из строкового буфера.
        strPaperName = Mid(String:=strPaperNamesList, _
            Start:=64 * (lngCounter - 1) + 1, Length:=64)
        intLength = VBA.InStr(Start:=1, String1:=strPaperName, String2:=Chr(0)) - 1
        strPaperName = Left(String:=strPaperName, Length:=intLength)
        
        ' Add a paper number and name to text string for the message box.
        ' Добавить номер бумаги и имя в строку сообщения
        strMsg = strMsg & vbCrLf & aintNumPaper(lngCounter) _
            & vbTab & strPaperName
            
    Next lngCounter
        
    ' Show the paper names in a message box.
    ' Вывести наименования бумаг в сообщении.
    MsgBox strMsg, , strMsgBox

GetPaperList_End:
    Exit Sub
    
GetPaperList_Err:
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
        Title:="Error Number " & Err.number & " Occurred"
    Resume GetPaperList_End
    
End Sub ' GetPaperList
' функция, которая выводит в сообщении список установленных
' в системе принтеров
Sub ShowPrinters()
    Dim strCount As String
    Dim strMsg As String
    Dim prtLoop As Printer
    
    On Error GoTo ShowPrinters_Err

    If Printers.Count > 0 Then
        ' Get count of installed printers.
        strMsg = "В системе установлены следующие принтеры: " & Printers.Count & vbCrLf & vbCrLf
    
        ' Enumerate printer system properties.
        For Each prtLoop In Application.Printers
            With prtLoop
                strMsg = strMsg _
                    & "Имя устройства: " & .DeviceName & vbCrLf _
                    & "Имя драйвера: " & .DriverName & vbCrLf _
                    & "Порт: " & .Port & vbCrLf & vbCrLf
            End With
        Next prtLoop
    
    Else
        strMsg = "Нет установленных принтеров."
    End If
    
    ' Display printer information.
    MsgBox Prompt:=strMsg, Buttons:=vbOKOnly, Title:=strMsgBox
    
ShowPrinters_End:
    Exit Sub
    
ShowPrinters_Err:
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
        Title:="Error Number " & Err.number & " Occurred"
    Resume ShowPrinters_End
    
End Sub ' ShowPrinters
' функция обработки диалога с пользователем по запросу печати
' на принтере, используемом по умолчанию
Function enabPrinting() As Boolean
Dim strDeviceName As String, Response As Integer
   
   strDeviceName = Application.Printer.DeviceName
   Response = MsgBox("Для печати будет использован принтер: " & _
      strDeviceName & "." & vbCrLf & "Печатать, используя этот принтер?", _
         vbExclamation + vbOKCancel, strMsgBox)
   If Response = 1 Then
      enabPrinting = True
   Else
      enabPrinting = False
   End If

End Function ' enabPrinting



