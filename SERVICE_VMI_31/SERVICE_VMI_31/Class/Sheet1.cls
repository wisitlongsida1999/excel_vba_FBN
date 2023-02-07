VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim target_position As String

Dim re As Object
Set re = New RegExp

'target_col = Mid(Target.Address, 2, 1)
'
'target_row = Mid(Target.Address, 4, 5)

re.Pattern = "([\d+\.]+)"
re.IgnoreCase = True
re.Global = True

t = Target.Address
x = re.Execute(t)


End Sub

