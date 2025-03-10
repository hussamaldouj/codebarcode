Attribute VB_Name = "Module2"
Sub SaveAsPDF()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String
    
    '  ⁄ÌÌ‰ «·Ê—ﬁ… Sheet21
Set ws = sheet21


    
    ' «·Õ’Ê· ⁄·Ï «·«”„ „‰ «·Œ·Ì… O10
    fileName = ws.Range("O10").Value
    
    ' «· Õﬁﬁ „‰ √‰ «·Œ·Ì… ·Ì”  ›«—€…
    If fileName = "" Then
        MsgBox "«·Œ·Ì… O10 ›«—€…° «·—Ã«¡ ≈œŒ«· «”„ «·„·›!", vbExclamation, "Œÿ√"
        Exit Sub
    End If
    
    ' «·”„«Õ ··„” Œœ„ »«Œ Ì«— „Ã·œ «·Õ›Ÿ
    With Application.FileDialog(msoFileDialogFolderPicker)
   '     .Title = "«Œ — „Ã·œ «·Õ›Ÿ"
   '     If .Show = -1 Then
   '         filePath = .SelectedItems(1) & "\" & fileName & ".pdf"
   filePath = sheet21.Range("O16").Value & fileName & ".pdf"

   '     Else
   '         MsgBox "·„ Ì „ «Œ Ì«— „Ã·œ!", vbExclamation, "≈·€«¡"
   '         Exit Sub
   '     End If
    End With
    
    ' ** ’œÌ— «·Ê—ﬁ… «·„Õœœ… ›ﬁÿ œÊ‰  €ÌÌ— «·Ê—ﬁ… «·‰‘ÿ…**
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End Sub

