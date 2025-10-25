VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAuthorizationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ласс, моделирующий объект, используемый дл€ хранени€ данных авторизованного пользовател€: '
'названи€ филиала, логина, парол€                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Dim reg As clsAuthorizationClass
'объ€вление локальных переменных дл€ хранени€ данных о филиале, логине, пароле
Dim filial As String, login As String, password As String, ID As Long

 'метод, сохран€ющий введенное название филиала в локальной переменной экземпл€ра класса
Public Sub setSubdiv(fl As String)

   filial = fl
   
End Sub      'setSubdiv
'метод, возвращающий при его вызове, сохраненное в локальной переменной экземпл€ра класса название филиала
Public Function getSubdiv() As String

   getSubdiv = filial
   
End Function 'getSubdiv
 'метод, сохран€ющий введенный логин в локальной переменной экземпл€ра класса
Public Sub setLogin(usLogin As String)

   login = usLogin
   
End Sub      'setLogin ()
'метод, возвращающий при его вызове, сохраненный в локальной переменной экземпл€ра класса логин
Public Function getLogin() As String

   getLogin = login
   
End Function 'getLogin()
 'метод, сохран€ющий введенный пароль в локальной переменной экземпл€ра класса
Public Sub setPassword(usPassword As String)

   password = usPassword
   
End Sub      'setPassword
'метод, возвращающий при его вызове, сохраненный в локальной переменной экземпл€ра класса пароль
Public Function getPassword() As String

   getPassword = password

End Function 'getPassword
'метод, сохран€ющий  од филиала в локальной переменной экземпл€ра класса
Public Sub setId(wId As Long)

   ID = wId
   
End Sub 'setId
'метод, возвращающий при его вызове, сохраненный в локальной переменной экземпл€ра класса  од филиала
Public Function getId() As Long

   getId = ID
   
End Function 'getId
