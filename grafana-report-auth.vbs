' =============================================================
' Script Name  : grafana_multi_ppt.vbs
' Description  : Downloads multiple images from Grafana and inserts them into a PowerPoint presentation.
' Author      : Pedro's Assistant
' =============================================================

Dim scriptPath, pptPath, apiKey
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
pptPath = scriptPath & "\grafana_dashboard.pptx"

' Grafana API Key (Replace with your key)
apiKey = "eyJrIjoiABC123..." ' ⚠️ Replace with your actual API Key

' Define Grafana image URLs (Modify as needed)
Dim imageUrls, imagePaths
imageUrls = Array( _
    "http://localhost:3000/render/d/Kdh0OoSGz2/panel1?orgId=1&format=png", _
    "http://localhost:3000/render/d/Kdh0OoSGz2/panel2?orgId=1&format=png", _
    "http://localhost:3000/render/d/Kdh0OoSGz2/panel3?orgId=1&format=png" _
)

' Download images and store file paths
ReDim imagePaths(UBound(imageUrls))
Dim i
For i = 0 To UBound(imageUrls)
    imagePaths(i) = scriptPath & "\grafana_" & (i+1) & ".png"
    If Not DownloadImage(imageUrls(i), imagePaths(i), apiKey) Then
        WScript.Echo "❌ Failed to download: " & imageUrls(i)
    End If
Next

' Insert images into PowerPoint
If InsertImagesIntoPPT(imagePaths, pptPath) Then
    WScript.Echo "✅ PowerPoint created successfully: " & pptPath
Else
    WScript.Echo "❌ Failed to create PowerPoint."
End If

' =============================================================
' Function: DownloadImage
' Purpose : Downloads an image from Grafana using an API Key.
' =============================================================
Function DownloadImage(url, filePath, apiKey)
    Dim objHTTP, objStream
    DownloadImage = False

    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "GET", url, False
    objHTTP.setRequestHeader "Authorization", "Bearer " & apiKey
    objHTTP.Send

    If objHTTP.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' Binary mode
        objStream.Open
        objStream.Write objHTTP.responseBody
        If objStream.Size > 0 Then
            objStream.SaveToFile filePath, 2 ' Overwrite if exists
            DownloadImage = True
        End If
        objStream.Close
        Set objStream = Nothing
    Else
        WScript.Echo "❌ HTTP Error: " & objHTTP.Status & " - " & objHTTP.StatusText
    End If

    Set objHTTP = Nothing
End Function

' =============================================================
' Function: InsertImagesIntoPPT
' Purpose : Creates a PowerPoint file and inserts multiple images in a grid.
' =============================================================
Function InsertImagesIntoPPT(imagePaths, pptPath)
    Dim pptApp, pptPres, pptSlide
    InsertImagesIntoPPT = False

    ' Create PowerPoint Application
    On Error Resume Next
    Set pptApp = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Then
        WScript.Echo "❌ PowerPoint is not installed."
        Exit Function
    End If
    On Error GoTo 0

    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add
    Set pptSlide = pptPres.Slides.Add(1, 1) ' ppLayoutTitle

    ' Grid Layout: 2 columns per row
    Dim imgWidth, imgHeight, startX, startY, colGap, rowGap
    imgWidth = 400 : imgHeight = 300 ' Adjust size
    startX = 50 : startY = 100
    colGap = 20 : rowGap = 30

    ' Insert images
    Dim row, col, posX, posY
    row = 0 : col = 0
    For i = 0 To UBound(imagePaths)
        posX = startX + (col * (imgWidth + colGap))
        posY = startY + (row * (imgHeight + rowGap))

        pptSlide.Shapes.AddPicture imagePaths(i), False, True, posX, posY, imgWidth, imgHeight

        ' Update row/column positions
        col = col + 1
        If col > 1 Then ' 2 columns per row
            col = 0
            row = row + 1
        End If
    Next

    ' Save and close
    pptPres.SaveAs pptPath
    pptPres.Close
    pptApp.Quit

    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    InsertImagesIntoPPT = True
End Function
