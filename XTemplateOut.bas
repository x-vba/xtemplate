Attribute VB_Name = "XTemplateOut"
Option Explicit

Private Function GetAllText() As String

    '@Description: This functions gathers all of the text in the various objects throughout the Outlook Mail or Appointment, including the To, Subject CC, BCC, and HTMLBody for the Mail, and To, Subject, and Body for the Appointment
    '@Author: Anthony Mancini
    '@License: MIT
    '@Version: 1.0.0
    '@Note: This function will differ for each Office program
    '@Returns: Returns a large string containing all of the text throughout the Document

    Dim allStrings As String
    
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


Private Function ParseOutTemplates( _
    ByVal allStrings As String) _
As Variant

    '@Description: This functions uses a Regex to parse out all the templates from the string provided. It also throws a few errors if it finds a poorly formatted template.
    '@Author: Anthony Mancini
    '@License: MIT
    '@Version: 1.0.0
    '@Note: This function will differ since Outlook will require absolute paths
    '@Param: allStrings is a string that will be regexed to find templates
    '@Returns: Returns a Dictionary in the following format: {OrigionalTemplate : FormattedTemplate}. The FormattedTemplate removes the curly braces and whitespace.

    ' Regexing out the templates
    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "\{\{.*?\}\}"
    End With
    
    Dim individualMatch As Variant
    Dim individualStringTemplate As String
    Dim regexMatches As Variant
    
    Set regexMatches = Regex.Execute(allStrings)
    
    ' Creating the dictionary that will be returned
    Dim templateDictionary As Object
    Set templateDictionary = CreateObject("Scripting.Dictionary")
    
    For Each individualMatch In regexMatches
        individualStringTemplate = individualMatch.Value
        individualStringTemplate = Mid(individualStringTemplate, 3, Len(individualStringTemplate) - 4)
        individualStringTemplate = Trim(individualStringTemplate)
        
        ' Checks if some of the templates are missing a curly brace, as if it
        ' finds 3 curly braces in a template it means one template is missing
        ' a brace
        If InStr(1, individualStringTemplate, "{") Or InStr(1, individualStringTemplate, "}") Then
            MsgBox "Error, missing curly brace '{' or '}' on one of the templates:" & vbCrLf & vbCrLf & individualMatch.Value, Title:="Template Syntax Error"
            Exit Function
        End If
        
        ' No check for full path necessary, as it will always be needed
        If Not templateDictionary.Exists(individualMatch.Value) Then
            templateDictionary.Add individualMatch.Value, individualStringTemplate
        End If
    Next
    
    Set ParseOutTemplates = templateDictionary

End Function


Private Function FetchExcelData( _
    ByVal templateDictionary As Variant) _
As Variant

    '@Description: This functions fetches out the data from the templates from the respective Excel files
    '@Author: Anthony Mancini
    '@License: MIT
    '@Version: 1.0.0
    '@Note: This function will be the same for each Office program
    '@Param: templateDictionary is a dictionary in the format: {OrigionalTemplate : FormattedTemplate}
    '@Returns: Returns a Dictionary in the following format: {OrigionalTemplate : FetchedValue}

    Dim ExcelApplication As Object
    Set ExcelApplication = CreateObject("Excel.Application")
    
    Dim currentWorkbook As Variant
    
    ExcelApplication.Visible = False

    
    Dim workbookPathDictionary As Object
    Set workbookPathDictionary = CreateObject("Scripting.Dictionary")

    Dim fetchTemplate As Variant
    Dim fullRangeDetails As String
    Dim workbookPath As String
    Dim workbookName As String
    Dim sheetName As String
    Dim rangeAddress As String
    
    ' Creating a workbook template dictionary containing collections
    ' of templates. This is used so that no workbook is opened up
    ' more than once when performing the fetches. The dictionary format
    ' is {WorkbookPath : templateDictionary}
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
    
    ' Actually performing the Excel Workbook fetches and creating a
    ' template dictionary in the following format:
    ' {Template : FetchedValue}
    Dim workbookPathKey As Variant
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
                        
            ' Perform the fetch
            If Not modifiedTemplateDictionary.Exists(fetchTemplate) Then
                Set currentWorkbook = ExcelApplication.Workbooks.Open(workbookPath)
                
                modifiedTemplateDictionary.Add fetchTemplate, currentWorkbook.Sheets(sheetName).Range(rangeAddress).Value
                
                currentWorkbook.Close False
                Set currentWorkbook = Nothing
            End If
        Next
    Next
    
    ' Replacing the other templates with the origional templates
    Dim templateKey As Variant
    
    For Each templateKey In templateDictionary.Keys()
        templateDictionary(templateKey) = modifiedTemplateDictionary(templateDictionary(templateKey))
    Next
    
    Set ExcelApplication = Nothing
    
    Set FetchExcelData = templateDictionary

End Function


Private Sub ReplaceTemplatesWithValues( _
    ByVal templateDictionary As Variant)

    '@Description: This subroutine replaces all the templates in the Outlook Mail or Appointment with their value
    '@Author: Anthony Mancini
    '@License: MIT
    '@Version: 1.0.0
    '@Note: This function will differ for each Office program
    '@Param: templateDictionary is a dictionary in the format: {OrigionalTemplate : FetchedValue}

    Dim templateKey As Variant
    
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

    '@Description: This subroutine performs all the steps to run XTemplate
    '@Author: Anthony Mancini
    '@License: MIT
    '@Version: 1.0.0
    '@Note: This function will be thes same for each Office program

    ' Getting all the strings
    Dim allStrings As String
    allStrings = GetAllText()
    
    ' Parsing out the templates
    Dim origionalTemplateDictionary As Variant
    Set origionalTemplateDictionary = ParseOutTemplates(allStrings)
        
    ' Fetching the data from Excel
    Dim templateDictionary As Variant
    Set templateDictionary = FetchExcelData(origionalTemplateDictionary)
    
    ' Replacing the templates with values
    ReplaceTemplatesWithValues templateDictionary

End Sub
