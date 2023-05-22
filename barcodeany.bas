Attribute VB_Name = "Module11"
' insert QR Code and Code128
' Author: Prasert Kanawattanachai
' prasert@cbs.chula.ac.th

Option Explicit

Public Sub qrcode_selection()
    insertBarcode qrcode(Sheet6.Range("v10").Value)
End Sub

Public Sub code128_selection()
    insertBarcode code128(Sheet6.Range("v10").Value)
End Sub

Public Function qrcode(data)
' API: https://barcode.tec-it.com/
    qrcode = "https://barcode.tec-it.com/barcode.ashx?code=QRCode&translate-esc=true&multiplebarcodes=false&unit=Fit&dpi=96&imagetype=png&eclevel=M&dmsize=Default&download=true&data=" & data
End Function

Public Function code128(data)
' API: https://barcode.tec-it.com/
    code128 = "https://barcode.tec-it.com/barcode.ashx?code=Code128&translate-esc=true&multiplebarcodes=false&unit=Fit&dpi=96&imagetype=Png&rotation=0&color=%23000000&bgcolor=%23ffffff&codepage=&qunit=Mm&quiet=0&data=" & data
End Function


'======================== PRIVATE FUNCTION ============================
Private Sub insertBarcode(barcode_url)
    'ActiveSheet.Pictures.Insert (barcode_url)
    Dim img_shape1 As Shape
    Set img_shape1 = ActiveSheet.Shapes.addPicture(Filename:=barcode_url, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=Range("J13:N16").Left + 2, _
        Top:=Range("J13:N16").Top + 2, _
        Width:=-1, Height:=-1)
    Dim img_shape2 As Shape
    Set img_shape2 = ActiveSheet.Shapes.addPicture(Filename:=barcode_url, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=Range("J32:N35").Left + 2, _
        Top:=Range("J32:N35").Top + 2, _
        Width:=-1, Height:=-1)
    Dim img_shape3 As Shape
    Set img_shape3 = ActiveSheet.Shapes.addPicture(Filename:=barcode_url, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=Range("J51:N54").Left + 2, _
        Top:=Range("J51:N54").Top + 2, _
        Width:=-1, Height:=-1)
        
End Sub
Sub deletepicture()
Dim pic As Picture
For Each pic In Sheet6.Pictures
If Not Application.Intersect(pic.TopLeftCell, Range("j13:j100")) Is Nothing Then
pic.Delete
End If
Next pic


End Sub
