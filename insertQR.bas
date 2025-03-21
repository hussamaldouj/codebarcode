Attribute VB_Name = "Module12"
Function Insert_QR(codetext As String)
    Dim URL As String, MyCell As Range

    Set MyCell = Application.Caller
    URL = "https://chart.googleapis.com/chart?chs=135x135&cht=qr&chl=" & codetext
    On Error Resume Next
      Sheet11.Pictures("My_QR_" & MyCell.Address(False, False)).Delete 'delete if there is prevoius one
    On Error GoTo 0
    Sheet11.Pictures.Insert(URL).Select
    With Selection.ShapeRange(1)
     .PictureFormat.CropLeft = 10
     .PictureFormat.CropRight = 10
     .PictureFormat.CropTop = 10
     .PictureFormat.CropBottom = 10
     .Name = "My_QR_" & MyCell.Address(False, False)
     .Left = MyCell.Left + 25
     .Top = MyCell.Top + 5
    End With
    Insert_QR = "" ' or some text to be displayed behind code
End Function
