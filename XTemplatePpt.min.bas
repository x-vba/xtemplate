Attribute VB_Name = "XTemplatePpt"
Option Explicit
Private Function GetAllText() As String
Dim individualSlide As Slide
Dim individualShape As Shape
Dim individualSmartArtNode As SmartArtNode
Dim individualRow As Row
Dim individualCell As Cell
Dim individualDesign As Design
Dim individualCustomLayout As CustomLayout
Dim allStrings$
For Each individualSlide In ActivePresentation.Slides
For Each individualShape In individualSlide.Shapes
On Error Resume Next
allStrings = allStrings + individualShape.TextFrame.TextRange.Text
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
allStrings = allStrings + individualSmartArtNode.TextFrame2.TextRange.Text
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
allStrings = allStrings + individualShape.Chart.ChartTitle.Text
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
allStrings = allStrings + individualCell.Shape.TextFrame.TextRange.Text
Next
Next
On Error GoTo 0
Next
On Error Resume Next
allStrings = allStrings + individualSlide.HeadersFooters.Header.Text
On Error GoTo 0
On Error Resume Next
allStrings = allStrings + individualSlide.HeadersFooters.Footer.Text
On Error GoTo 0
Next
For Each individualDesign In ActivePresentation.Designs
For Each individualShape In individualDesign.SlideMaster.Shapes
On Error Resume Next
allStrings = allStrings + individualShape.TextFrame.TextRange.Text
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
allStrings = allStrings + individualSmartArtNode.TextFrame2.TextRange.Text
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
allStrings = allStrings + individualShape.Chart.ChartTitle.Text
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
allStrings = allStrings + individualCell.Shape.TextFrame.TextRange.Text
Next
Next
On Error GoTo 0
Next
For Each individualCustomLayout In individualDesign.SlideMaster.CustomLayouts
For Each individualShape In individualCustomLayout.Shapes
On Error Resume Next
allStrings = allStrings + individualShape.TextFrame.TextRange.Text
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
allStrings = allStrings + individualSmartArtNode.TextFrame2.TextRange.Text
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
allStrings = allStrings + individualShape.Chart.ChartTitle.Text
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
allStrings = allStrings + individualCell.Shape.TextFrame.TextRange.Text
Next
Next
On Error GoTo 0
Next
Next
Next
GetAllText = allStrings
End Function
Private Function ParseOutTemplates( ByVal allStrings$)
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "\{\{.*?\}\}"
End With
Dim individualMatch
Dim individualStringTemplate$
Dim regexMatches
Set regexMatches = Regex.Execute(allStrings)
Dim templateDictionary As Object
Set templateDictionary = CreateObject("Scripting.Dictionary")
For Each individualMatch In regexMatches
individualStringTemplate = individualMatch.Value
individualStringTemplate = Mid(individualStringTemplate, 3, Len(individualStringTemplate) - 4)
individualStringTemplate = Trim(individualStringTemplate)
If InStr(1, individualStringTemplate, "{") Or InStr(1, individualStringTemplate, "}") Then
MsgBox "Error, missing curly brace '{' or '}' on one of the templates:" & vbCrLf & vbCrLf & individualMatch.Value, Title:="Template Syntax Error"
Exit Function
End If
If InStr(1, individualStringTemplate, "\") Then
If Not templateDictionary.Exists(individualMatch.Value) Then
templateDictionary.Add individualMatch.Value, individualStringTemplate
End If
Else
If Not templateDictionary.Exists(individualMatch.Value) Then
templateDictionary.Add individualMatch.Value, ActivePresentation.Path & "\" & individualStringTemplate
End If
End If
Next
Set ParseOutTemplates = templateDictionary
End Function
Private Function FetchExcelData( ByVal templateDictionary)
Dim ExcelApplication As Object
Set ExcelApplication = CreateObject("Excel.Application")
Dim currentWorkbook
ExcelApplication.Visible = False
Dim workbookPathDictionary As Object
Set workbookPathDictionary = CreateObject("Scripting.Dictionary")
Dim fetchTemplate
Dim fullRangeDetails$
Dim workbookPath$
Dim workbookName$
Dim sheetName$
Dim rangeAddress$
For Each fetchTemplate In templateDictionary.Keys()
fullRangeDetails = Right(templateDictionary(fetchTemplate), Len(templateDictionary(fetchTemplate)) - InStrRev(templateDictionary(fetchTemplate), "\"))
workbookName = Left(fullRangeDetails, InStrRev(fullRangeDetails, "]") - 1)
workbookName = Mid(workbookName, 2)
workbookPath = Left(templateDictionary(fetchTemplate), InStrRev(templateDictionary(fetchTemplate), "\")) & workbookName
If Not workbookPathDictionary.Exists(workbookPath) Then
workbookPathDictionary.Add workbookPath, New Collection
workbookPathDictionary.Item(workbookPath).Add templateDictionary(fetchTemplate)
Else
workbookPathDictionary.Item(workbookPath).Add templateDictionary(fetchTemplate)
End If
Next
Dim workbookPathKey
Dim modifiedTemplateDictionary As Object
Set modifiedTemplateDictionary = CreateObject("Scripting.Dictionary")
For Each workbookPathKey In workbookPathDictionary.Keys()
For Each fetchTemplate In workbookPathDictionary(workbookPathKey)
fullRangeDetails = Right(fetchTemplate, Len(fetchTemplate) - InStrRev(fetchTemplate, "\"))
workbookName = Left(fullRangeDetails, InStrRev(fullRangeDetails, "]") - 1)
workbookName = Mid(workbookName, 2)
workbookPath = Left(fetchTemplate, InStrRev(fetchTemplate, "\")) & workbookName
sheetName = Mid(fullRangeDetails, InStrRev(fullRangeDetails, "]") + 1)
sheetName = Left(sheetName, InStrRev(sheetName, "!") - 1)
rangeAddress = Right(fullRangeDetails, Len(fullRangeDetails) - InStrRev(fullRangeDetails, "!"))
rangeAddress = Replace(rangeAddress, "$", "")
If Not modifiedTemplateDictionary.Exists(fetchTemplate) Then
Set currentWorkbook = ExcelApplication.Workbooks.Open(workbookPath)
modifiedTemplateDictionary.Add fetchTemplate, currentWorkbook.Sheets(sheetName).Range(rangeAddress).Value
currentWorkbook.Close False
Set currentWorkbook = Nothing
End If
Next
Next
Dim templateKey
For Each templateKey In templateDictionary.Keys()
templateDictionary(templateKey) = modifiedTemplateDictionary(templateDictionary(templateKey))
Next
Set ExcelApplication = Nothing
Set FetchExcelData = templateDictionary
End Function
Private Sub ReplaceTemplatesWithValues( ByVal templateDictionary)
Dim individualSlide As Slide
Dim individualShape As Shape
Dim individualSmartArtNode As SmartArtNode
Dim individualRow As Row
Dim individualCell As Cell
Dim individualDesign As Design
Dim individualCustomLayout As CustomLayout
Dim templateKey
For Each templateKey In templateDictionary.Keys()
For Each individualSlide In ActivePresentation.Slides
For Each individualShape In individualSlide.Shapes
On Error Resume Next
individualShape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
individualSmartArtNode.TextFrame2.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
individualShape.Chart.ChartTitle.Text = Replace(individualShape.Chart.ChartTitle.Text, templateKey, templateDictionary(templateKey))
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
individualCell.Shape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
Next
On Error GoTo 0
Next
On Error Resume Next
individualSlide.HeadersFooters.Header.Text = Replace(individualSlide.HeadersFooters.Header.Text, templateKey, templateDictionary(templateKey))
On Error GoTo 0
On Error Resume Next
individualSlide.HeadersFooters.Footer.Text = Replace(individualSlide.HeadersFooters.Footer.Text, templateKey, templateDictionary(templateKey))
On Error GoTo 0
Next
For Each individualDesign In ActivePresentation.Designs
For Each individualShape In individualDesign.SlideMaster.Shapes
On Error Resume Next
individualShape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
individualSmartArtNode.TextFrame2.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
individualShape.Chart.ChartTitle.Text = Replace(individualShape.Chart.ChartTitle.Text, templateKey, templateDictionary(templateKey))
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
individualCell.Shape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
Next
On Error GoTo 0
Next
For Each individualCustomLayout In individualDesign.SlideMaster.CustomLayouts
For Each individualShape In individualCustomLayout.Shapes
On Error Resume Next
individualShape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
On Error GoTo 0
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
individualSmartArtNode.TextFrame2.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
End If
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
individualShape.Chart.ChartTitle.Text = Replace(individualShape.Chart.ChartTitle.Text, templateKey, templateDictionary(templateKey))
End If
End If
On Error Resume Next
For Each individualRow In individualShape.Table.Rows
For Each individualCell In individualRow.Cells
individualCell.Shape.TextFrame.TextRange.Replace templateKey, templateDictionary(templateKey)
Next
Next
On Error GoTo 0
Next
Next
Next
Next
End Sub
Public Sub XTemplate()
Dim allStrings$
allStrings = GetAllText()
Dim origionalTemplateDictionary
Set origionalTemplateDictionary = ParseOutTemplates(allStrings)
Dim templateDictionary
Set templateDictionary = FetchExcelData(origionalTemplateDictionary)
ReplaceTemplatesWithValues templateDictionary
End Sub