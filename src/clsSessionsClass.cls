VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSessionsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ласс, моделирующий объект, используемый дл€ отметки штампов времени начала и конца сессий   '
'работы с базой данных, а также идентификации пользовател€ каждой сессии                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Dim sess As clsSessionsClass

'метод, сохран€ющий врем€/дату начала запущенной сессии и аттрибуты идентификации пользовател€
Public Sub setSessionBegin(subDiv As String, login As String)

   CurrentDb.Execute " INSERT INTO sessions ( subdivision,user_name,state, host_name ) VALUES " _
                      & "('" & subDiv & "','" & login & "','открыта','" & GetThisComputerName & " ' );"
   
End Sub 'setSessionBegin
'метод, сохран€ющий врем€/дату начала запущенной сессии и аттрибуты идентификации пользовател€
Public Sub setSessionEnd(subDiv As String, login As String)

   CurrentDb.Execute " INSERT INTO sessions ( subdivision,user_name,state, host_name ) VALUES " _
                      & " ('" & subDiv & "','" & login & "','закрыта', '" & GetThisComputerName & " ' );"
                      
End Sub 'setSessionEnd



