Attribute VB_Name = "Module4"
Function GetLastFourDigits(studentID As String) As String
    Dim lastFour As String
    
    ' ������� ��� 4 �����
    lastFour = Right(studentID, 4)
    
    ' ��� ��� ����� ��� �� 1000� ����� ����� �����
    If Val(lastFour) <= 999 Then
        lastFour = CStr(Val(lastFour)) ' ����� ����� ��� �� ��� ����� ����� �����
    End If
    
    GetLastFourDigits = lastFour
End Function

Sub InsertStudentImage()
    Dim ws As Worksheet
    Dim studentID As String
    Dim imagePath As String
    Dim pictureShape As Shape
    Dim imageFolder As String
    
    ' ����� ������ Sheet21
    Set ws = sheet21
    
    ' ����� ���� ���� ����� (�� ������ ������ ��� ���� ������ ����)
    imageFolder = ws.Range("O18").Value
    
    ' ������ ��� ����� ������ �� I6
    studentID = ws.Range("O10").Value
    
    ' ������� ��� 3 ����� ���
    finalID = GetLastFourDigits(studentID)
    
    ' ����� ������ ������ ������
    If Dir(imageFolder & finalID & ".PNG") <> "" Then
        imagePath = imageFolder & finalID & ".PNG"
    ElseIf Dir(imageFolder & finalID & ".png") <> "" Then
        imagePath = imageFolder & finalID & ".png"
    Else
        MsgBox "������ ��� ������ ������ ��� " & finalID, vbExclamation, "���"
        Exit Sub
    End If
    
    ' ��� �� ���� ����� ������ ������
    For Each pictureShape In ws.Shapes
        If pictureShape.Name = "StudentImage" Then
            pictureShape.Delete
            Exit For
        End If
    Next pictureShape
    If Dir(imagePath) = "" Then
        MsgBox "������ ��� ������ ��: " & imagePath, vbExclamation, "���"
        Exit Sub
    End If

    ' ����� ������
    Set pictureShape = ws.Shapes.AddPicture( _
    fileName:=imagePath, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=ws.Range("I6").Left, _
    Top:=ws.Range("I6").Top + 10, _
    Width:=100, _
    Height:=120)

End Sub

