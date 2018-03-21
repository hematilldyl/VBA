VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

===============================================================
'SubmitSurvey
'---------------------------------------------------------------
' Author: Dylan Hematillake
' Version: 1.5
' Purpose: Once the survey sheet has been filled, hides the sheets
'          and only shows a blank sheet, print the results in the 
'          report template as a pdf and then emails the pdf to
'          the data collector
'===============================================================
Private Sub CommandButton1_Click()


 
'Workbook variables
Dim ws As Worksheet
Dim wb As Workbook
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Results")
Set ws2 = wb.Sheets("Summary Report")

'Outlook variables
Dim objOutlook As Outlook.Application
Dim objOutlookMsg As Outlook.MailItem
Dim objOutlookRecip As Outlook.recipient
Dim objOutlookAttach As Outlook.Attachment
Dim Pathtofile, Pathtofile2 As String




'Activate the hidden summary report
Worksheets("Summary Report").Visible = True
Worksheets("Summary Report").Activate

'iterate through the score column in Results and the two columns
'in Summary Report

For i = 5 To 9
    For j = 1 To 2
    
    If j = 1 Then Set r1 = ws.Range("V" & i + 5)
    If j = 2 Then Set r1 = ws.Range("V" & i + 10)
    If j = 1 Then Set r2 = ws2.Range("A" & i)
    If j = 2 Then Set r2 = ws2.Range("F" & i)
    
    If r1 >= 0.8 Then r2.Interior.Color = RGB(185, 247, 183)
    
    If r1 >= 0.6 And r1 < 0.8 Then r2.Interior.Color = RGB(183, 183, 247)
    
    If r1 <= 0.6 Then r2.Interior.Color = RGB(247, 183, 183)
    Next j
Next i

'ThisWorkbook.Sheets(Array("Survey", "Summary Report")).Select
'Print the results to pdf, will display results but not responses.
Pathtofile = CStr(Application.ActiveWorkbook.Path & "\Results.pdf")
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Pathtofile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
ThisWorkbook.Sheets("Survey").Select
Pathtofile2 = CStr(Application.ActiveWorkbook.Path & "\Responses.pdf")
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Pathtofile2, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False

'Set up email variables
Set objOutlook = CreateObject("Outlook.Application")
Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

'Send the email to Recipients (can be multi)
'Subject = subject header, Body = body text of email, Attachments.Add to select the file (as path)

With objOutlookMsg
    Set objOutlookRecip = .Recipients.Add("dylan.hematillake@value-infinity.com")
    objOutlookRecip.Type = olTo
   .Subject = "Automated Response to Talent Management Survey"
   .Body = "This is an automated response to a completed Talent Management Survey"
   .Attachments.Add Pathtofile
   .Attachments.Add Pathtofile2
   
   For Each objOutlookRecip In .Recipients
       objOutlookRecip.Resolve
   Next
   If DisplayMsg Then
       .Display
   Else
   'Change to .send if you wish for only a draft to be saved and not sent
       .Send
   End If
End With
Set objOutlook = Nothing

'Excel sheet termination, first disable warnings then delete
Application.DisplayAlerts = False
wb.Sheets("Sheet1").Visible = xlSheetVisible
wb.Sheets("Intro").Delete
wb.Sheets("Survey").Delete
wb.Sheets("Model").Delete
wb.Sheets("Summary Report").Delete
wb.Sheets("Results").Delete

'Autosaves to prevent user interruption
ActiveWorkbook.Save
                            
End Sub
