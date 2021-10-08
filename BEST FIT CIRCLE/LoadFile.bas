Attribute VB_Name = "LoadFile"
Sub loadF()
    'On Error GoTo lp1:
    Form1.CommonDialog1.FileName = ""
    Form1.CommonDialog1.DialogTitle = "OPEN FILE"
    Form1.CommonDialog1.Filter = "Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
    Form1.CommonDialog1.Action = 1
    Open Form1.CommonDialog1.FileName For Input Access Read As #1
    Data = 1
    Do While Not EOF(1)
        ReDim Preserve X(Data)
        ReDim Preserve Y(Data)
        Input #1, X(Data), Y(Data)
        Data = Data + 1
     Loop
     Close #1
    N = Data - 1
lp1:
End Sub

 
