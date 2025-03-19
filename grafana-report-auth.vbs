' =============================================================
' Script Name  : grafana_to_ppt.vbs
' Description  : Downloads a Grafana panel as an image with authentication and inserts it into a PowerPoint presentation.
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
DownloadImage "http://localhost:3000/render/d/Kdh0OoSGz2/windows-exporter-dashboard-2024-v2?orgId=1", imagePath, username, password
CreatePowerPoint imagePath, pptPath

' =============================================================
' Function: DownloadImage
' Purpose : Downloads an image from the specified URL with authentication.
' =============================================================
Sub DownloadImage(url, filePath, user, pass)
    Dim objHTTP, objStream, auth
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    ' Encode credentials for Basic Authentication
    auth = "Basic " & Base64Encode(user & ":" & pass)

    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "Authorization", auth
    objHTTP.Send

    ' Check if the request was successful
    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binary mode
        objStream.Open
        objStream.Write objHTTP.responseBody
        objStream.SaveToFile filePath, 2 ' Overwrite if the file exists
        objStream.Close
        Set objStream = Nothing
    Else
        WScript.Echo "❌ Failed to download image. HTTP Status: " & objHTTP.Status
        WScript.Quit
    End If

    Set objHTTP = Nothing
End Sub

' =============================================================
' Function: Base64Encode
' Purpose : Encodes a string into Base64 (needed for Basic Authentication)
' =============================================================
Function Base64Encode(str)
    Dim objXML, objNode
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("Base64")
    objNode.DataType = "bin.base64"
    objNode.NodeTypedValue = Stream_StringToBinary(str)
    Base64Encode = objNode.Text
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' =============================================================
' Function: Stream_StringToBinary
' Purpose : Converts a string to binary for Base64 encoding
' =============================================================
Function Stream_StringToBinary(text)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' Text
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText text
    objStream.Position = 0
    objStream.Type = 1 ' Binary
    Stream_StringToBinary = objStream.Read
    objStream.Close
    Set objStream = Nothing
End Function

' =============================================================
' Function: CreatePowerPoint
' Purpose : Creates a PowerPoint presentation and inserts the downloaded image.
' =============================================================
Sub CreatePowerPoint(imagePath, outputPath)
    Dim pptApp, pptPresentation, slide, image
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    Set pptPresentation = pptApp.Presentations.Add
    Set slide = pptPresentation.Slides.Add(1, 1)
    slide.Shapes.Title.TextFrame.TextRange.Text = "Grafana Dashboard Panel"

    Set image = slide.Shapes.AddPicture(imagePath, False, True, 50, 100, 600, 400)

    pptPresentation.SaveAs outputPath
    pptPresentation.Close
    pptApp.Quit
End Sub
