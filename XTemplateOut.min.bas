Attribute VB_Name = "XTemplateOut"
Option Explicit
Private Function GetAllText() As String
Dim allStrings$
If TypeName(Application.ActiveInspector.CurrentItem) = "MailItem" Then
Dim individualMailItem As MailItem
Set individualMailItem = Application.ActiveInspector.CurrentItem
allStrings = allStrings + individualMailItem.To
allStrings = allStrings + individualMailItem.Subject
allStrings = allStrings + individualMailItem.CC
allStrings = allStrings + individualMailItem.BCC
allStrings = allStrings + individualMailItem.HTMLBody
ElseIf TypeName(Application.ActiveInspector.CurrentItem) = "AppointmentItem" Then
Dim individualAppointmentItem As AppointmentItem
Set individualAppointmentItem = Application.ActiveInspector.CurrentItem
allStrings = allStrings + individualAppointmentItem.Subject
allStrings = allStrings + individualAppointmentItem.Location
allStrings = allStrings + individualAppointmentItem.Body
End If
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
If Not templateDictionary.Exists(individualMatch.Value) Then
templateDictionary.Add individualMatch.Value, individualStringTemplate
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
Dim templateKey
If TypeName(Application.ActiveInspector.CurrentItem) = "MailItem" Then
Dim individualMailItem As MailItem
Set individualMailItem = Application.ActiveInspector.CurrentItem
For Each templateKey In templateDictionary.Keys()
individualMailItem.To = Replace(individualMailItem.To, templateKey, templateDictionary(templateKey))
individualMailItem.Subject = Replace(individualMailItem.Subject, templateKey, templateDictionary(templateKey))
individualMailItem.CC = Replace(individualMailItem.CC, templateKey, templateDictionary(templateKey))
individualMailItem.BCC = Replace(individualMailItem.BCC, templateKey, templateDictionary(templateKey))
individualMailItem.HTMLBody = Replace(individualMailItem.HTMLBody, templateKey, templateDictionary(templateKey))
Next
ElseIf TypeName(Application.ActiveInspector.CurrentItem) = "AppointmentItem" Then
Dim individualAppointmentItem As AppointmentItem
Set individualAppointmentItem = Application.ActiveInspector.CurrentItem
For Each templateKey In templateDictionary.Keys()
individualAppointmentItem.Subject = Replace(individualAppointmentItem.Subject, templateKey, templateDictionary(templateKey))
individualAppointmentItem.Location = Replace(individualAppointmentItem.Location, templateKey, templateDictionary(templateKey))
individualAppointmentItem.Body = Replace(individualAppointmentItem.Body, templateKey, templateDictionary(templateKey))
Next
End If
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