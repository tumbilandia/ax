' =============================================================
' Script Name  : weekly_grafana_report.vbs
' Description  : Generates a PowerPoint report for a Grafana panel.
'               - Asks for a start date.
'               - Calculates the corresponding week number.
'               - Names the report as "rpt80-wXX-service.pptx".
'               - Checks if the report already exists and asks for deletion.
' Author       : Pedro Alvarado
' Date         : YYYY-MM-DD
' =============================================================

Dim startDate, weekNumber, reportPath
Dim scriptPath, imagePath, pptName

' Get script path to store files
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' 1. Ask for start date
startDate = InputBox("Enter the start date (yyyy/mm/dd):", "Start Date")

' Validate input format
If Not IsDate(startDate) Then
    MsgBox "Invalid date format. Please enter a valid date (yyyy/mm/dd).", vbCritical, "Error"
    WScript.Quit
End If

' 2. Calculate week number
weekNumber = DatePart("ww", CDate(startDate), vbSunday, vbFirstFourDays)

' 3. Generate report name
pptName = "rpt80-w" & weekNumber & "-service.pptx"
reportPath = scriptPath & "\" & pptName
imagePath = scriptPath & "\grafana_panel.png"

' 4. Check if report already exists
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(reportPath) Then
    Dim response
    response = MsgBox("The report '" & pptName & "' already exists. Do you want to delete it?", vbYesNo + vbQuestion, "File Exists")
    If response = vbYes Then
        fso.DeleteFile reportPath, True
    Else
        MsgBox "Exiting without changes.", vbInformation, "Exit"
        WScript.Quit
    End If
End If

' Call functions to download the image and create PowerPoint
DownloadImage "http://localhost:3000/d/Kdh0OoSGz2/windows-exporter-dashboard-2024-v2?orgId=1&from=2025-03-20T03:29:41.954Z&to=2025-03-20T06:29:41.954Z&timezone=browser&var-job=windows&var-hostname=$__all&var-instance=localhost:9182&var-show_hostname=mango&viewPanel=panel-49", imagePath
CreatePowerPoint imagePath, reportPath

' =============================================================
' Function: DownloadImage
' Purpose : Downloads an image from the specified URL and saves it locally.
' =============================================================
Sub DownloadImage(url, filePath)
    Dim objHTTP, objStream
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "GET", url, False
    objHTTP.Send

    ' Check if request was successful
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binary mode
        objStream.Open
        objStream.Write objHTTP.responseBody
        objStream.SaveToFile filePath, 2 ' Overwrite
        objStream.Close
        Set objStream = Nothing
    Else
        MsgBox "❌ Failed to download image. HTTP Status: " & objHTTP.Status, vbCritical, "Error"
        WScript.Quit
    End If

    Set objHTTP = Nothing
End Sub

' =============================================================
' Function: CreatePowerPoint
' Purpose : Creates a PowerPoint presentation and inserts the downloaded image.
' =============================================================
Sub CreatePowerPoint(imagePath, outputPath)
    Dim pptApp, pptPresentation, slide, image

    ' Create PowerPoint instance
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Create new presentation
    Set pptPresentation = pptApp.Presentations.Add

    ' Add slide
    Set slide = pptPresentation.Slides.Add(1, 1)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Grafana Weekly Report - W" & weekNumber

    ' Insert image
    Set image = slide.Shapes.AddPicture(imagePath, False, True, 50, 100, 600, 400)

    ' Save and close
    pptPresentation.SaveAs outputPath
    pptPresentation.Close
    pptApp.Quit
End Sub
