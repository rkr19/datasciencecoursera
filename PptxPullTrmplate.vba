Sub ExtractTemplateFromPPTX()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim templatePath As String
    
    ' Path to the PowerPoint file
    Dim pptFilePath As String
    pptFilePath = "C:\Path\To\Your\File.pptx" ' Update with your file path
    
    ' Create a new instance of PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = False ' Set to True if you want to see PowerPoint
    
    ' Open the PowerPoint file
    Set pptPresentation = pptApp.Presentations.Open(pptFilePath)
    
    ' Get the path to the current template
    templatePath = pptPresentation.Designs(1).TemplateName
    
    ' Save the template to a specific folder
    pptPresentation.ApplyTemplate templatePath
    pptPresentation.SaveAs "C:\Path\To\Save\Template\" & templatePath & ".potx" ' Update with your save path
    
    ' Close the PowerPoint presentation without saving changes
    pptPresentation.Close False
    
    ' Clean up
    pptApp.Quit
    Set pptPresentation = Nothing
    Set pptApp = Nothing
    
    MsgBox "Template extracted successfully!"
End Sub