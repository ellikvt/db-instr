Attribute VB_Name = "main"
Option Compare Database
Option Explicit

'��������� ������ clsAuthorizationClass
Public auth As Object
'��������� ������ clsSessionsClass
Public session As Object
'��������� ������ clsRelocateClass
Public rl As Object
' ��������� ������ clsSelectRecordClass
Public sr As Object
Public cancelUpdateQwr As Boolean
'������� ���������� �������
Public recordsCount As Long
'���������� ��� ������ ������� ��� ����������� ���
Public queryString As String
' ��������� ��������� (������� � ��������� ���� MsgBox)
Public Const strMsgBox As String = "���������� ���"
' ��� ������� ������_�������, ��������������� ���� F27 �� ��������� "������� � ���� ������"
Public Const SOURCE_ID = 147
' ��� ������� ����_��������, ��������������� ���� F27 �� ��������� "����� ��������"
Public Const DESTINATION_ID = 226

'���������� ������ ������� ��������� ����� ���������� ���������� Win API
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
   
   
   
'������������� ���������� ����������
Public Sub initVarsInMain()

   cancelUpdateQwr = False
   recordsCount = 0
   queryString = ""
   
End Sub 'initVarsInMain

'��������� ����� ���������� ���������� � ������� ������ Windows API
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
      GetThisComputerName = "���������� �������� ���."
   End If

End Function 'GetThisComputerName

'����������� �������
'���������� ������, ������������ �������� WHERE ������� ���
'� �������� �����, ��� ���������� � ������ ������������� ����
Public Function qwrParamsStr(value As Long) As String
Dim rst As Recordset

   Set rst = CurrentDb.OpenRecordset("SELECT ���,parent_id FROM ��������� " _
                                    & "WHERE parent_id = " & value & "")
   If rst.RecordCount <> 0 Then
      '���� ������������ ������
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
      '�������� ��� ��������
      qwrParamsStr = "structure_id = " & CStr(value)
   End If
   
   rst.Close
   
End Function 'qwrParamsStr
'���������� ��� �� ������� ���������, ��������������� ���� �������, ��������� ���� �������� ������� � ��������
'��������� value. ���� � �������� ��������� ������� ��� �������� (������), �� ������������ ��� ���
Public Function getDesiredSubdivCode(value As Long) As Long
Dim rst As Recordset, Code As Long, condition As Boolean

   Set rst = CurrentDb.OpenRecordset("SELECT ���,parent_id,subdivision FROM ��������� " _
                                    & "WHERE ��� = " & value & "")
                                    
   Do Until rst.Fields(2) = True
      Set rst = CurrentDb.OpenRecordset("SELECT ���,parent_id,subdivision FROM ��������� " _
                                    & "WHERE ��� = " & rst.Fields(1) & "")
   Loop
   getDesiredSubdivCode = rst.Fields(0)
      
   rst.Close
   
End Function 'getDesiredSubdivCode
' �������, �������:
' a. ���������� ������� Windows API DeviceCapabilities ��� ������ ���������
' �. ��������� ��������� �� ��������� �. �) � �������� �� ���������
' �. ��������� ��������� �� ��������� �. �) � ������ �������������� ����� ������
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
    
    ' �������� ��� � ���� �������� �� ���������.
    strDeviceName = Application.Printer.DeviceName
    strDevicePort = Application.Printer.Port
    
    ' Get the count of paper names supported by the printer.
    ' �������� ���������� ������������ ������, �������������� ���������
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERNAMES, _
        lpOutput:=ByVal vbNullString, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Re-dimension the array to the count of paper names.
    ' ������ ������ ������� ��� ������������ ������
    ReDim aintNumPaper(1 To lngPaperCount)
    
    ' Pad the variable to accept 64 bytes for each paper name.
    ' ����������� ���������� 64 ���� ��� �������� ������� ������������ ������
    strPaperNamesList = String(64 * lngPaperCount, 0)

    ' Get the string buffer of all paper names supported by the printer.
    ' �������� ��������� ����� ��� ���� ������������ ������, �������������� ���������
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERNAMES, _
        lpOutput:=ByVal strPaperNamesList, _
        lpDevMode:=DEFAULT_VALUES)
    
    ' Get the array of all paper numbers supported by the printer.
    ' �������� ������ ���� ������� ������������, �������������� ���������.
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _
        lpPort:=strDevicePort, _
        iIndex:=DC_PAPERS, _
        lpOutput:=aintNumPaper(1), _
        lpDevMode:=DEFAULT_VALUES)
    
    ' List the available paper names.
    ' ������� ��������� ������������ ������.
    strMsg = "� ��������� ������ ������������ �������: " & strDeviceName & _
      "." & vbCrLf & "��� ���� �������� ��������� ���� ������:" & vbCrLf
    For lngCounter = 1 To lngPaperCount
        
        ' Parse a paper name from the string buffer.
        ' �������� ������������ ������ �� ���������� ������.
        strPaperName = Mid(String:=strPaperNamesList, _
            Start:=64 * (lngCounter - 1) + 1, Length:=64)
        intLength = VBA.InStr(Start:=1, String1:=strPaperName, String2:=Chr(0)) - 1
        strPaperName = Left(String:=strPaperName, Length:=intLength)
        
        ' Add a paper number and name to text string for the message box.
        ' �������� ����� ������ � ��� � ������ ���������
        strMsg = strMsg & vbCrLf & aintNumPaper(lngCounter) _
            & vbTab & strPaperName
            
    Next lngCounter
        
    ' Show the paper names in a message box.
    ' ������� ������������ ����� � ���������.
    MsgBox strMsg, , strMsgBox

GetPaperList_End:
    Exit Sub
    
GetPaperList_Err:
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical & vbOKOnly, _
        Title:="Error Number " & Err.number & " Occurred"
    Resume GetPaperList_End
    
End Sub ' GetPaperList
' �������, ������� ������� � ��������� ������ �������������
' � ������� ���������
Sub ShowPrinters()
    Dim strCount As String
    Dim strMsg As String
    Dim prtLoop As Printer
    
    On Error GoTo ShowPrinters_Err

    If Printers.Count > 0 Then
        ' Get count of installed printers.
        strMsg = "� ������� ����������� ��������� ��������: " & Printers.Count & vbCrLf & vbCrLf
    
        ' Enumerate printer system properties.
        For Each prtLoop In Application.Printers
            With prtLoop
                strMsg = strMsg _
                    & "��� ����������: " & .DeviceName & vbCrLf _
                    & "��� ��������: " & .DriverName & vbCrLf _
                    & "����: " & .Port & vbCrLf & vbCrLf
            End With
        Next prtLoop
    
    Else
        strMsg = "��� ������������� ���������."
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
' ������� ��������� ������� � ������������� �� ������� ������
' �� ��������, ������������ �� ���������
Function enabPrinting() As Boolean
Dim strDeviceName As String, Response As Integer
   
   strDeviceName = Application.Printer.DeviceName
   Response = MsgBox("��� ������ ����� ����������� �������: " & _
      strDeviceName & "." & vbCrLf & "��������, ��������� ���� �������?", _
         vbExclamation + vbOKCancel, strMsgBox)
   If Response = 1 Then
      enabPrinting = True
   Else
      enabPrinting = False
   End If

End Function ' enabPrinting



