' VBScript to generate a PowerPoint presentation with multiple images
Dim objPPT, objPres, objSlide
Dim slideWidth, slideHeight
Dim textContent, imgFolder
Dim framePaths(10), i

' Set the image folder path (Modify this path accordingly)
imgFolder = "d:\_dev\sopra\projects\airbus\01-OpenShift\80-monitoring\img\"

' Set text content for frame 1
textContent = "This is the text of slide 1"

' Define image paths for frames 2-10
For i = 2 To 10
    'framePaths(i) = imgFolder & "frame0" & i & ".png"
    framePaths(i) = imgFolder & "frame" & right(100+i,2) & ".png"
Next

' Create PowerPoint Application
Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True ' Make it visible

' Create Presentation
Set objPres = objPPT.Presentations.Add()

' Add a new slide (Title Slide)
Set objSlide = objPres.Slides.Add(1, 1) ' 1 = ppLayoutTitle

' Get slide size (PowerPoint default: 960x540 points)
slideWidth = objPres.PageSetup.SlideWidth
slideHeight = objPres.PageSetup.SlideHeight

' Define frame positions and sizes (adjust as needed)
objSlide.Shapes.AddTextbox(1, 10, 10, slideWidth - 20, 50).TextFrame.TextRange.Text = textContent ' Frame 1 (Text)

' Insert images into predefined positions
objSlide.Shapes.AddPicture framePaths(2), False, True, 10, 70, 220, 150 ' Frame 2
objSlide.Shapes.AddPicture framePaths(3), False, True, 240, 70, 220, 150 ' Frame 3
objSlide.Shapes.AddPicture framePaths(4), False, True, 470, 70, 220, 150 ' Frame 4

objSlide.Shapes.AddPicture framePaths(5), False, True, 10, 230, 100, 100 ' Frame 5
objSlide.Shapes.AddPicture framePaths(6), False, True, 120, 230, 100, 100 ' Frame 6
objSlide.Shapes.AddPicture framePaths(7), False, True, 230, 230, 220, 100 ' Frame 7
objSlide.Shapes.AddPicture framePaths(8), False, True, 460, 230, 100, 100 ' Frame 8
objSlide.Shapes.AddPicture framePaths(9), False, True, 570, 230, 100, 100 ' Frame 9

objSlide.Shapes.AddPicture framePaths(10), False, True, 10, 340, slideWidth - 20, 150 ' Frame 10

' Save the presentation
objPres.SaveAs imgFolder & "presentation.pptx"

' Cleanup
Set objSlide = Nothing
Set objPres = Nothing
Set objPPT = Nothing

MsgBox "PowerPoint presentation created successfully!", vbInformation, "Done"
