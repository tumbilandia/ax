' VBScript to generate a PowerPoint presentation with a fixed layout
Dim objPPT, objPres, objSlide
Dim objShape, objImage
Dim slideWidth, slideHeight
Dim imagePath
Dim outputFileName

' Set output file name
outputFileName = "mock_presentation.pptx"

' Define image path (Replace with actual image path)
imagePath = "d:\_dev\sopra\projects\airbus\01-OpenShift\80-monitoring\img\02.png"

' Create PowerPoint application
Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True ' Make PowerPoint visible

' Create a new presentation
Set objPres = objPPT.Presentations.Add

' Define slide size
slideWidth = objPres.PageSetup.SlideWidth
slideHeight = objPres.PageSetup.SlideHeight

' Add a new slide
Set objSlide = objPres.Slides.Add(1, 1) ' 1 = ppLayoutText

' Define frame positions (adjust as needed)
Dim frames(10,3)

' Frame 1 (Text Box)
frames(1,0) = 20 ' Left
frames(1,1) = 10 ' Top
frames(1,2) = slideWidth - 40 ' Width
frames(1,3) = 50 ' Height

' Frame 2 - 10 (Images)
frames(2,0) = 20: frames(2,1) = 70: frames(2,2) = 250: frames(2,3) = 200
frames(3,0) = 280: frames(3,1) = 70: frames(3,2) = 250: frames(3,3) = 200
frames(4,0) = 540: frames(4,1) = 70: frames(4,2) = 250: frames(4,3) = 200
frames(5,0) = 20: frames(5,1) = 280: frames(5,2) = 120: frames(5,3) = 100
frames(6,0) = 150: frames(6,1) = 280: frames(6,2) = 120: frames(6,3) = 100
frames(7,0) = 280: frames(7,1) = 280: frames(7,2) = 250: frames(7,3) = 100
frames(8,0) = 540: frames(8,1) = 280: frames(8,2) = 120: frames(8,3) = 100
frames(9,0) = 670: frames(9,1) = 280: frames(9,2) = 120: frames(9,3) = 100
frames(10,0) = 20: frames(10,1) = 400: frames(10,2) = slideWidth - 40: frames(10,3) = 150

' Add text to frame 1
Set objShape = objSlide.Shapes.AddTextbox(1, frames(1,0), frames(1,1), frames(1,2), frames(1,3))
objShape.TextFrame.TextRange.Text = "This is the text of slide 1"
objShape.TextFrame.TextRange.Font.Size = 20

' Add images to frames 2 - 10
For i = 2 To 10
    Set objImage = objSlide.Shapes.AddPicture(imagePath, False, True, frames(i,0), frames(i,1), frames(i,2), frames(i,3))
Next

' Save the presentation
objPres.SaveAs CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".\") & "\" & outputFileName
objPres.Close
objPPT.Quit

' Clean up
Set objShape = Nothing
Set objImage = Nothing
Set objSlide = Nothing
Set objPres = Nothing
Set objPPT = Nothing

WScript.Echo "PowerPoint file created: " & outputFileName
