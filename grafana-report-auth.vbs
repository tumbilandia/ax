' =============================================================
' Script Name  : grafana_to_ppt.vbs
' Description  : Downloads an image from Grafana using an API Key and inserts it into a PowerPoint slide.
' =============================================================

' Set output file paths
Dim scriptPath, imagePath, pptPath
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
imagePath = scriptPath & "\grafana_panel.png"
pptPath = scriptPath & "\grafana_presentation.pptx"

' Grafana API Key (replace with your actual key)
Dim apiKey
apiKey = "eyJrIjoiABC123..."  ' ⚠️ Replace this with your actual API Key

' Image URL
Dim imageUrl
imageUrl = "http://localhost:3000/render/d/Kdh0OoSGz2/windows-exporter-dashboard-2024-v2?orgId=1&format=png"

' Download the image
If DownloadImage(imageUrl, imagePath, apiKey) Then
    WScript.Echo "✅ Image downloaded successfully: " & imagePath
    ' Insert into PowerPoint
    If InsertImageIntoPPT(imagePath, pptPath) Then
        WScript.Echo "✅ PowerPoint created successfully: " & pptPath
    Else
        WScript.Echo "❌ Failed to create PowerPoint."
    End If
Else
    WScript.Echo "❌ Failed to download image."
End If

' =============================================================
' Function: DownloadImage
' Purpose : Downloads an image from Grafana using an API Key.
' =============================================================
Function DownloadImage(url, filePath, apiKey)
    Dim objHTTP, objStream
    DownloadImage = False ' Default to failure

    ' Create HTTP request
    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "Authorization", "Bearer " & apiKey
    objHTTP.Send

    ' Check if request was successful
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binary mode
        objStream.Open
        objStream.Write objHTTP.responseBody

        ' Ensure the file is an image
        If objStream.Size > 0 Then
            objStream.SaveToFile filePath, 2 ' Overwrite if exists
            DownloadImage = True
        Else
            WScript.Echo "❌ Image file is empty. Check authentication or URL."
        End If

        objStream.Close
        Set objStream = Nothing
    Else
        WScript.Echo "❌ HTTP Error: " & objHTTP.Status & " - " & objHTTP.StatusText
    End If

    Set objHTTP = Nothing
End Function

' =============================================================
' Function: InsertImageIntoPPT
' Purpose : Creates a PowerPoint file and inserts the downloaded image.
' =============================================================
Function InsertImageIntoPPT(imagePath, pptPath)
    Dim pptApp, pptPres, pptSlide
    InsertImageIntoPPT = False ' Default to failure

    ' Create PowerPoint Application
    On Error Resume Next
    Set pptApp = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Then
        WScript.Echo "❌ PowerPoint is not installed."
        Exit Function
    End If
    On Error GoTo 0

    pptApp.Visible = True ' Show PowerPoint

    ' Create a new presentation
    Set pptPres = pptApp.Presentations.Add
    Set pptSlide = pptPres.Slides.Add(1, 1) ' 1 = ppLayoutTitle

    ' Insert image
    On Error Resume Next
    pptSlide.Shapes.AddPicture imagePath, False, True, 100, 100, 600, 400
    If Err.Number <> 0 Then
        WScript.Echo "❌ Failed to insert image into PowerPoint."
        Exit Function
    End If
    On Error GoTo 0

    ' Save and close
    pptPres.SaveAs pptPath
    pptPres.Close
    pptApp.Quit

    ' Cleanup
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    InsertImageIntoPPT = True
End Function
