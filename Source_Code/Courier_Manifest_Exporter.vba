Sub ExportCourierManifests()
    Dim wb As Workbook
    Dim wsMain As Worksheet
    Dim wsTemp As Worksheet
    Dim courierNames As Variant
    Dim sheetNames As Variant
    Dim i As Integer
    Dim folderPath As String
    Dim fileName As String

    Set wsMain = ThisWorkbook.Sheets("Logistics_Data")
    courierNames = Array("Aramex", "EMX", "DHL")
    sheetNames = Array("Aramex_Format", "EMX_Format", "DHL_Format")
    
    folderPath = ThisWorkbook.Path & "\"

    For i = LBound(courierNames) To UBound(courierNames)
        On Error Resume Next
        Set wsTemp = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        
        If Not wsTemp Is Nothing Then
            
            wsTemp.Copy
            Set wb = ActiveWorkbook
            
            fileName = courierNames(i) & "_Manifest_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
            
            Application.DisplayAlerts = False
            wb.SaveAs folderPath & fileName
            wb.Close SaveChanges:=False
            Application.DisplayAlerts = True
            
            MsgBox courierNames(i) & " file has been saved successfully!", vbInformation
        Else
            MsgBox sheetNames(i) & " sheet not found!", vbExclamation
        End If
    Next i
End Sub