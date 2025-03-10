Attribute VB_Name = "Module2"
Sub SaveAsPDF()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileName As String
    
    ' ����� ������ Sheet21
Set ws = sheet21


    
    ' ������ ��� ����� �� ������ O10
    fileName = ws.Range("O10").Value
    
    ' ������ �� �� ������ ���� �����
    If fileName = "" Then
        MsgBox "������ O10 ����ɡ ������ ����� ��� �����!", vbExclamation, "���"
        Exit Sub
    End If
    
    ' ������ �������� ������� ���� �����
    With Application.FileDialog(msoFileDialogFolderPicker)
   '     .Title = "���� ���� �����"
   '     If .Show = -1 Then
   '         filePath = .SelectedItems(1) & "\" & fileName & ".pdf"
   filePath = sheet21.Range("O16").Value & fileName & ".pdf"

   '     Else
   '         MsgBox "�� ��� ������ ����!", vbExclamation, "�����"
   '         Exit Sub
   '     End If
    End With
    
    ' **����� ������ ������� ��� ��� ����� ������ ������**
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End Sub

