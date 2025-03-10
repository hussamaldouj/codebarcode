Attribute VB_Name = "Module4"
Function GetLastFourDigits(studentID As String) As String
    Dim lastFour As String
    
    ' «” Œ—«Ã ¬Œ— 4 √—ﬁ«„
    lastFour = Right(studentID, 4)
    
    ' ≈–« ﬂ«‰ «·—ﬁ„ √ﬁ· „‰ 1000° ≈“«·… «·’›— «·√Ê·
    If Val(lastFour) <= 999 Then
        lastFour = CStr(Val(lastFour)) '  ÕÊÌ· «·—ﬁ„ ≈·Ï ‰’ »⁄œ ≈“«·… «·’›— «·√Ê·
    End If
    
    GetLastFourDigits = lastFour
End Function

Sub InsertStudentImage()
    Dim ws As Worksheet
    Dim studentID As String
    Dim imagePath As String
    Dim pictureShape As Shape
    Dim imageFolder As String
    
    '  ÕœÌœ «·Ê—ﬁ… Sheet21
    Set ws = sheet21
    
    '  ÕœÌœ „”«— „Ã·œ «·’Ê— (ﬁ„ » €ÌÌ— «·„”«— Õ”» „Êﬁ⁄ «·„Ã·œ ·œÌﬂ)
    imageFolder = ws.Range("O18").Value
    
    ' «·Õ’Ê· ⁄·Ï «·—ﬁ„ «·√ŒÌ— „‰ I6
    studentID = ws.Range("O10").Value
    
    ' «” Œ—«Ã ¬Œ— 3 √—ﬁ«„ ›ﬁÿ
    finalID = GetLastFourDigits(studentID)
    
    '  ÕœÌœ «·„”«— «·ﬂ«„· ··’Ê—…
    If Dir(imageFolder & finalID & ".PNG") <> "" Then
        imagePath = imageFolder & finalID & ".PNG"
    ElseIf Dir(imageFolder & finalID & ".png") <> "" Then
        imagePath = imageFolder & finalID & ".png"
    Else
        MsgBox "«·’Ê—… €Ì— „ÊÃÊœ… ··ÿ«·» —ﬁ„ " & finalID, vbExclamation, "Œÿ√"
        Exit Sub
    End If
    
    ' Õ–› √Ì ’Ê—… ﬁœÌ„… „ÊÃÊœ… „”»ﬁ«
    For Each pictureShape In ws.Shapes
        If pictureShape.Name = "StudentImage" Then
            pictureShape.Delete
            Exit For
        End If
    Next pictureShape
    If Dir(imagePath) = "" Then
        MsgBox "«·’Ê—… €Ì— „ÊÃÊœ… ›Ì: " & imagePath, vbExclamation, "Œÿ√"
        Exit Sub
    End If

    ' ≈œ—«Ã «·’Ê—…
    Set pictureShape = ws.Shapes.AddPicture( _
    fileName:=imagePath, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=ws.Range("I6").Left, _
    Top:=ws.Range("I6").Top + 10, _
    Width:=100, _
    Height:=120)

End Sub

