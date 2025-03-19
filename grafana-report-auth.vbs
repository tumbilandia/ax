' =============================================================
' Script Name  : grafana_to_ppt.vbs
' Description  : Downloads a Grafana panel image and inserts it into PowerPoint.
' =============================================================

' Define script path to store files in the same directory where the script runs
Dim scriptPath, imagePath, pptPath
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
imagePath = scriptPath & "\grafana_panel.png"
pptPath = scriptPath & "\grafana_dashboard.pptx"

' Grafana credentials
Dim username, password
username = "admin"  ' Change this to your Grafana username
password = "admin"  ' Change this to your Grafana password

' Call functions to download image and create PowerPoint
If DownloadImage("http://localhost:3000/render/d/Kdh0OoSGz2/windows-exporter-dashboard-2024-v2?orgId=1", imagePath, username, password) Then
    CreatePowerPoint imagePath, pptPath
Else
    WScript.Echo "❌ Failed to download image. Exiting..."
End If

' =============================================================
' Function: DownloadImage
' Purpose : Downloads an image from the specified URL with authentication.
' =============================================================
Function DownloadImage(url, filePath, user, pass)
    Dim objHTTP, objStream, base64Auth
    DownloadImage = False ' Default to failure

    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Encode credentials in Base64 using a correct method
    base64Auth = "Basic " & EncodeBase64(user & ":" & pass)

    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "Authorization", base64Auth
    objHTTP.Send

    ' Check if the request was successful
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binary mode
        objStream.Open
        objStream.Write objHTTP.responseBody

        ' Ensure the file is an image
        If objStream.Size > 0 Then
            objStream.SaveToFile filePath, 2 ' Overwrite if the file exists
            DownloadImage = True ' Success
        Else
            WScript.Echo "❌ Image file is empty. Possible authentication or URL issue."
        End If

        objStream.Close
        Set objStream = Nothing
    Else
        WScript.Echo "❌ HTTP Error: " & objHTTP.Status & " - " & objHTTP.StatusText
    End If

    Set objHTTP = Nothing
End Function

' =============================================================
' Function: EncodeBase64
' Purpose : Correctly encodes a string in Base64 for HTTP Basic Authentication.
' =============================================================
Function EncodeBase64(text)
    Dim bytes, objXML, objNode
    Set bytes = CreateObject("ADODB.Stream")
    bytes.Type = 2 ' Text mode
    bytes.Charset = "utf-8"
    bytes.Open
    bytes.WriteText text
    bytes.Position = 0
    bytes.Type = 1 ' Convert to binary

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = bytes.Read
    EncodeBase64 = Replace(objNode.Text, vbLf, "")

    Set objNode = Nothing
    Set objXML = Nothing
    Set bytes = Nothing
End Function

' =============================================================
' Function: CreatePowerPoint
' Purpose : Creates a PowerPoint presentation and inserts the downloaded image.
' =============================================================
Sub CreatePowerPoint(imagePath, outputPath)
    Dim pptApp, pptPresentation, slide, image

    ' Check if the file exists before inserting
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(imagePath) Then
        WScript.Echo "❌ Image file not found: " & imagePath
        Exit Sub
    End If
    Set fso = Nothing

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    Set pptPresentation = pptApp.Presentations.Add
    Set slide = pptPresentation.Slides.Add(1, 1)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Grafana Dashboard Panel"

    ' Insert the image
    Set image = slide.Shapes.AddPicture(imagePath, False, True, 50, 100, 600, 400)

    ' Save and close
    pptPresentation.SaveAs outputPath
    pptPresentation.Close
    pptApp.Quit

    WScript.Echo "✅ PowerPoint created successfully: " & outputPath
End Sub
