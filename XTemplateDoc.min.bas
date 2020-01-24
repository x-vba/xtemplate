Attribute VB_Name = "XTemplateDoc"
Option Explicit
Private Function GetAllText() As String
Dim individualShape As Shape
Dim individualInlineShape As InlineShape
Dim individualSmartArtNode As SmartArtNode
Dim individualSection As Section
Dim individualHeaderFooter As HeaderFooter
Dim allStrings$
allStrings = ActiveDocument.Content.Text
For Each individualShape In ActiveDocument.Shapes
allStrings = allStrings + individualShape.TextFrame.TextRange.Text
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
allStrings = allStrings + individualInlineShape.Range.Text
Next
For Each individualShape In ActiveDocument.Shapes
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
allStrings = allStrings + individualSmartArtNode.TextFrame2.TextRange.Text
Next
End If
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
If individualInlineShape.HasSmartArt Then
For Each individualSmartArtNode In individualInlineShape.SmartArt.AllNodes
allStrings = allStrings + individualSmartArtNode.TextFrame2.TextRange.Text
Next
End If
Next
For Each individualSection In ActiveDocument.Sections
For Each individualHeaderFooter In individualSection.Headers
allStrings = allStrings + individualHeaderFooter.Range.Text
Next
For Each individualHeaderFooter In individualSection.Footers
allStrings = allStrings + individualHeaderFooter.Range.Text
Next
Next
For Each individualShape In ActiveDocument.Shapes
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
allStrings = allStrings + individualShape.Chart.ChartTitle.Text
End If
End If
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
If individualInlineShape.HasChart Then
If individualInlineShape.Chart.HasTitle Then
allStrings = allStrings + individualInlineShape.Chart.ChartTitle.Text
End If
End If
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
If InStr(1, individualStringTemplate, Application.PathSeparator) Then
If Not templateDictionary.Exists(individualMatch.Value) Then
templateDictionary.Add individualMatch.Value, individualStringTemplate
End If
Else
If Not templateDictionary.Exists(individualMatch.Value) Then
templateDictionary.Add individualMatch.Value, ActiveDocument.Path & Application.PathSeparator & individualStringTemplate
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
fullRangeDetails = Right(templateDictionary(fetchTemplate), Len(templateDictionary(fetchTemplate)) - InStrRev(templateDictionary(fetchTemplate), Application.PathSeparator))
workbookName = Left(fullRangeDetails, InStrRev(fullRangeDetails, "]") - 1)
workbookName = Mid(workbookName, 2)
workbookPath = Left(templateDictionary(fetchTemplate), InStrRev(templateDictionary(fetchTemplate), Application.PathSeparator)) & workbookName
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
fullRangeDetails = Right(fetchTemplate, Len(fetchTemplate) - InStrRev(fetchTemplate, Application.PathSeparator))
workbookName = Left(fullRangeDetails, InStrRev(fullRangeDetails, "]") - 1)
workbookName = Mid(workbookName, 2)
workbookPath = Left(fetchTemplate, InStrRev(fetchTemplate, Application.PathSeparator)) & workbookName
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
Dim individualShape As Shape
Dim individualInlineShape As InlineShape
Dim individualSmartArtNode As SmartArtNode
Dim individualSection As Section
Dim individualHeaderFooter As HeaderFooter
Dim templateKey
Dim modifiedTemplateKey$
For Each templateKey In templateDictionary.Keys()
With ActiveDocument.Range.Find
.Text = templateKey
.Replacement.Text = templateDictionary(templateKey)
.Execute Replace:=wdReplaceAll
End With
For Each individualShape In ActiveDocument.Shapes
individualShape.TextFrame.TextRange.Text = Replace(individualShape.TextFrame.TextRange.Text, templateKey, templateDictionary(templateKey))
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
With individualInlineShape.Range.Find
.Text = templateKey
.Replacement.Text = templateDictionary(templateKey)
.Execute Replace:=wdReplaceAll
End With
Next
For Each individualShape In ActiveDocument.Shapes
If individualShape.HasSmartArt Then
For Each individualSmartArtNode In individualShape.SmartArt.AllNodes
individualSmartArtNode.TextFrame2.TextRange.Text = Replace(individualSmartArtNode.TextFrame2.TextRange.Text, templateKey, templateDictionary(templateKey))
Next
End If
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
If individualInlineShape.HasSmartArt Then
For Each individualSmartArtNode In individualInlineShape.SmartArt.AllNodes
individualSmartArtNode.TextFrame2.TextRange.Text = Replace(individualSmartArtNode.TextFrame2.TextRange.Text, templateKey, templateDictionary(templateKey))
Next
End If
Next
For Each individualSection In ActiveDocument.Sections
For Each individualHeaderFooter In individualSection.Headers
With individualHeaderFooter.Range.Find
.Text = templateKey
.Replacement.Text = templateDictionary(templateKey)
.Execute Replace:=wdReplaceAll
End With
Next
For Each individualHeaderFooter In individualSection.Footers
With individualHeaderFooter.Range.Find
.Text = templateKey
.Replacement.Text = templateDictionary(templateKey)
.Execute Replace:=wdReplaceAll
End With
Next
Next
For Each individualShape In ActiveDocument.Shapes
If individualShape.HasChart Then
If individualShape.Chart.HasTitle Then
individualShape.Chart.ChartTitle.Text = Replace(individualShape.Chart.ChartTitle.Text, templateKey, templateDictionary(templateKey))
End If
End If
Next
For Each individualInlineShape In ActiveDocument.InlineShapes
If individualInlineShape.HasChart Then
If individualInlineShape.Chart.HasTitle Then
individualInlineShape.Chart.ChartTitle.Text = Replace(individualInlineShape.Chart.ChartTitle.Text, templateKey, templateDictionary(templateKey))
End If
End If
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