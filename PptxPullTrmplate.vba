Sub ExtractTemplateFromPPTX()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim design As Object
    Dim templatePath As String
    
    ' Path to the PowerPoint file
    Dim pptFilePath As String
    pptFilePath = "C:\Path\To\Your\File.pptx" ' Update with your file path
    
    ' Create a new instance of PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = False ' Set to True if you want to see PowerPoint
    
    ' Open the PowerPoint file
    Set pptPresentation = pptApp.Presentations.Open(pptFilePath)
    
    ' Loop through each design in the presentation
    For Each design In pptPresentation.Designs
        ' Get the path to the template
        templatePath = "C:\Path\To\Save\Template\" & design.Name & ".potx" ' Update with your save path
        
        ' Save the template
        pptPresentation.ApplyTemplate design.Name
        pptPresentation.SaveAs templatePath
    Next design
    
    ' Close the PowerPoint presentation without saving changes
    pptPresentation.Close False
    
    ' Clean up
    pptApp.Quit
    Set pptPresentation = Nothing
    Set pptApp = Nothing
    
    MsgBox "Templates extracted successfully!"
End Sub