VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTreeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Класс описывающий объект TreeView и операции с ним '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit
 
'Объявляем класс Tree с событиями
Public WithEvents Tree As TreeView
Attribute Tree.VB_VarHelpID = -1
Public Tbl As String
Public fldParent As String
Public fldKey As String
Public fldText As String
Public createKey As Long
Dim nod As Node


Private Sub Class_Initialize()
'Инициализируем переменные класса для работы с таблицей
   Tbl = "Tbl"
   fldParent = "Parent"
   fldKey = "Key"
   fldText = "Text"
End Sub

' События при управлении левой кнопкой мыши
Private Sub Tree_Click()
    ' MsgBox Tree.SelectedItem.Key
End Sub

'Добавление основного узла
Public Sub AddBaseNode(Key As String, Text As String)
Dim idx As Long

   Tree.Nodes.Add , , Key, Text
   'idx = Tree.Nodes.Add(, , Key).Index
   'With Tree.Nodes(idx)
   '   .Text = Text
   'End With
End Sub

'Добавление дочернего узла
Public Sub AddNode(Parent As String, Key As String, Text As String)
Dim idx As Long

   Tree.Nodes.Add Parent, tvwChild, Key, Text
   'Tree.Nodes.Add , , Key, Text
    'idx = Tree.Nodes.Add(Parent, tvwChild, Key).Index
    'With Tree.Nodes(idx)
    '   .Text = Text
    'End With
End Sub

'Очистка дерева
Public Sub ClearNode()
    Tree.Nodes.Clear
End Sub

Public Sub GenerateRecursive(Parent As String)
Dim r As DAO.Recordset
Dim Key As String
Dim par As String
Dim Text As String
'========================================================================'
'                РЕКУРСИВНАЯ ГЕНЕРАЦИЯ ДЕРЕВА                    '
'========================================================================'
Set r = CurrentDb.OpenRecordset("SELECT * FROM " & Tbl & _
    " WHERE " & fldParent & "=" & Parent & ";", dbOpenDynaset)
If r.EOF And r.BOF Then
Else
    r.MoveFirst
    While Not r.EOF
        Key = "key" & r.Fields(fldKey)
        par = "key" & r.Fields(fldParent)
        Text = r.Fields(fldText)
        If r.Fields(fldParent) = 0 Then
            AddBaseNode Key, Text
        Else
            AddNode par, Key, Text
        End If
        GenerateRecursive r.Fields(fldKey)
        r.MoveNext
    Wend
End If
'========================================================================
r.Close
Set r = Nothing
End Sub

'Генерация дерева из таблицы
Public Sub GenerateTree()
Dim r As DAO.Recordset
Dim Key As String
Dim par As String
Dim Text As String

ClearNode

GenerateRecursive "0"

End Sub

'Получить код элемента
Public Function GetKey() As Long
    GetKey = DelKeyStr(Tree.SelectedItem.Key)
End Function

'Удалить префикс
Private Function DelKeyStr(Text As String) As Long
Dim stroka As String
    stroka = Right(Text, Len(Text) - 3)
DelKeyStr = CLng(stroka)
End Function

'Добавить ветку
Public Sub AddTblNode(Parent As String, Text As String)
Dim Key As String
Dim par As String
Dim LastId As Long

CurrentDb.Execute "INSERT INTO " & Tbl & " ( [" & fldText & "], " & fldParent & _
" ) SELECT """ & Text & """ AS Txt, " & DelKeyStr(Parent) & " AS Prn;"
LastId = DMax(fldKey, Tbl, "")

createKey = LastId
Key = "key" & LastId
If DelKeyStr(Parent) = 0 Then
    AddBaseNode Key, Text
Else
    AddNode Parent, Key, Text
End If
End Sub

'Обновить ветку
Public Sub UpdateTblNode(Key As String, UpdText As String)
CurrentDb.Execute "UPDATE " & Tbl & " SET " & fldText & "=""" & UpdText & """ WHERE " & fldKey & "=" & _
DelKeyStr(Key) & ";"
Tree.Nodes.Item(Key).Text = UpdText
End Sub

'Удалить ветку
Public Sub DelTblNode(Key As String)
CurrentDb.Execute "DELETE * FROM " & Tbl & " WHERE " & fldKey & "=" & _
DelKeyStr(Key) & ";"
Tree.Nodes.Remove Key
End Sub

'Рекурсивное удаление ветки (если есть дочерние и внучатые ветки)
Public Sub RecursiveDelTblNode(Key As String)
Dim r As Recordset

Set r = CurrentDb.OpenRecordset("SELECT * FROM " & Tbl & _
    " WHERE " & fldParent & "=" & DelKeyStr(Key) & ";", dbOpenDynaset)
If r.EOF And r.BOF Then
    CurrentDb.Execute "DELETE * FROM " & Tbl & " WHERE " & _
        fldKey & "=" & DelKeyStr(Key) & ";"
    Tree.Nodes.Remove Key
Else
    r.MoveFirst
    While Not r.EOF
        RecursiveDelTblNode "key" & r.Fields(fldKey)
        r.MoveNext
    Wend
    CurrentDb.Execute "DELETE * FROM " & Tbl & " WHERE " & _
        fldKey & "=" & DelKeyStr(Key) & ";"
    Tree.Nodes.Remove Key
End If

r.Close
Set r = Nothing
End Sub



