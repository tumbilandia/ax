' =============================================================
' Script Name  : download_image.vbs
' Description  : Downloads an image from Grafana using an API Key.
' =============================================================

' Set output file path
Dim scriptPath, imagePath
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
imagePath = scriptPath & "\grafana_panel.png"

' Grafana API Key (replace with your actual key)
Dim apiKey
apiKey = "eyJrIjoiABC123..."  ' ⚠️ Replace this with your API Key

' Image URL
Dim imageUrl
imageUrl = "http://localhost:3000/render/d/Kdh0OoSGz2/windows-exporter-dashboard-2024-v2?orgId=1&format=png"

' Download the image
If DownloadImage(imageUrl, imagePath, apiKey) Then
    WScript.Echo "✅ Image downloaded successfully: " & imagePath
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
