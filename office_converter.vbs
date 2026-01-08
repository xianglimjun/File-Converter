Option Explicit

Dim fso, inputFile, outputFile, ext
Dim wordApp, doc
Dim pptApp, pres

If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: office_converter.vbs <input> <output>"
    WScript.Quit 1
End If

inputFile = WScript.Arguments(0)
outputFile = WScript.Arguments(1)

Set fso = CreateObject("Scripting.FileSystemObject")
inputFile = fso.GetAbsolutePathName(inputFile)
outputFile = fso.GetAbsolutePathName(outputFile)
ext = LCase(fso.GetExtensionName(inputFile))

On Error Resume Next

If ext = "docx" Or ext = "doc" Then
    Set wordApp = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error creating Word object: " & Err.Description
        WScript.Quit 1
    End If

    wordApp.Visible = False
    wordApp.DisplayAlerts = 0 ' wdAlertsNone
    
    Set doc = wordApp.Documents.Open(inputFile)
    If Err.Number <> 0 Then
        WScript.Echo "Error opening document: " & Err.Description
        wordApp.Quit
        WScript.Quit 1
    End If

    ' 17 = wdFormatPDF
    doc.SaveAs outputFile, 17
    If Err.Number <> 0 Then
        WScript.Echo "Error saving as PDF: " & Err.Description
        doc.Close 0
        wordApp.Quit
        WScript.Quit 1
    End If

    doc.Close 0 ' wdDoNotSaveChanges
    wordApp.Quit

ElseIf ext = "pptx" Or ext = "ppt" Then
    Set pptApp = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error creating PowerPoint object: " & Err.Description
        WScript.Quit 1
    End If

    ' Keep PowerPoint hidden - use MsoTriState values
    pptApp.DisplayAlerts = 1 ' ppAlertsNone
    
    ' Open(FileName, ReadOnly, Untitled, WithWindow=0 means no window)
    Set pres = pptApp.Presentations.Open(inputFile, -1, 0, 0)
    If Err.Number <> 0 Then
        WScript.Echo "Error opening presentation: " & Err.Description
        pptApp.Quit
        WScript.Quit 1
    End If

    ' 32 = ppSaveAsPDF
    pres.SaveAs outputFile, 32
    If Err.Number <> 0 Then
        WScript.Echo "Error saving as PDF: " & Err.Description
        pres.Close
        pptApp.Quit
        WScript.Quit 1
    End If

    pres.Close
    pptApp.Quit

Else
    WScript.Echo "Unsupported format: " & ext
    WScript.Quit 1
End If

If Err.Number <> 0 Then
    WScript.Echo "Unhandled Error: " & Err.Description
    WScript.Quit 1
End If

WScript.Quit 0
