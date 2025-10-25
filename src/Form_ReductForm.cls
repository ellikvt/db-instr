VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ReductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'нажатие кнопки "Запрос приборов во всех узлах"
Private Sub allQueryButton_Click()
Dim rst As Recordset

   Me.RecordSource = "SELECT * FROM Склад "
   statusLabel.Caption = "ТМЦ во всех узлах структуры"
   '"популяризация" источника данных формы
   Set rst = Me.Recordset
   rst.MoveLast
   rst.MoveFirst

End Sub 'allQueryButton_Click

'загрузка формы
Private Sub Form_Load()
Dim rst As Recordset
Dim meRst As Recordset
Dim Noth As Variant
            
   'инициализация нового экземпляра класса, описывающего виды перемещений ТМЦ
   Set rl = New clsRelocateClass
         
   '"популяризация" источника данных формы
   Set meRst = Me.Recordset
   meRst.MoveLast
   meRst.MoveFirst
         
End Sub 'Form_Load


'обнуление элементов массива, хранящего путь
Private Sub arrayInitilize(wpathArray() As String)
Dim i As Integer

   For i = 0 To 14
      wpathArray(i) = ""
   Next i
   
End Sub 'arrayInitilize

'нажатие кнопки Печать отчета ЗИП АСУ
Private Sub printAsuReportButton_Click()
   
   DoCmd.OpenReport "AsuReport", acViewNormal

End Sub 'printAsuReportButton_Click

'переместить текущее ТМЦ в пределах структуры
Private Sub relocateButton_Click()

   'метод Inside - перемещение структура - структура
   rl.Inside
   rl.inRelocation = True
   DoCmd.OpenForm "UpdateСклад", , , , , acDialog

End Sub 'relocateButton_Click
'нажатие кнопки перемещения ТМЦ из структуры в буфер приборов
Private Sub relocateToBufButton_Click()

   'метод InsideOute - перемещение структура - буфер
   rl.InsideOut
   rl.insOutRelocation = True
   DoCmd.OpenForm "UpdateСклад", , , , , acDialog

End Sub 'relocateToBufButton_Click

'нажатие кнопки перемещения ТМЦ из буфера приборов в структуру
Private Sub relocateUnnestedButton_Click()

End Sub 'relocateUnnestedButton_Click

'обновить содержимое буфера приборов
Private Sub unnestedQueryButton_Click()
   
   
End Sub 'unnestedQueryButton_Click(

