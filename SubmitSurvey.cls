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
' Purpose: Once the survey sheet has been filled, hides the sheets
'          and only shows a blank sheet, print the results in the 
'          report template as a pdf and then emails the pdf to
'          the data collector
'===============================================================
Private Sub CommandButton1_Click()

Dim ws As Worksheet
Dim wb As Workbook

Dim objOutlook As Outlook.Application
Dim objOutlookMsg As Outlook.MailItem
Dim objOutlookRecip As Outlook.recipient
Dim objOutlookAttach As Outlook.Attachment
Dim Pathtofile As String

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Results")
Set ws2 = wb.Sheets("Summary Report")

ws2.Visible = True
ws2.Activate

For i = 5 To 9
    For j = 1 To 2
    
    If j = 1 Then Set r1 = ws.Range("V" & i + 5)
    If j = 2 Then Set r1 = ws.Range("V" & i + 10)
    If j = 1 Then Set r2 = ws2.Range("A" & i)
    If j = 2 Then Set r2 = ws2.Range("F" & i)
    
    If r1 >= 0.8 Then r2.Interior.Color = RGB(0, 106, 255)
    
    If r1 >= 0.6 Then r2.Interior.Color = RGB(100, 100, 240)
    
    If r1 <= 0.6 Then r2.Interior.Color = RGB(139, 178, 240)
    Next j
Next i

Pathtofile = CStr(Application.ActiveWorkbook.Path & "\Results.pdf")
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Pathtofile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
        



Set objOutlook = CreateObject("Outlook.Application")

Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

With objOutlookMsg
    Set objOutlookRecip = .Recipients.Add("dylan.hematillake@the_email.com")
    objOutlookRecip.Type = olTo
   .Subject = "Test"
   .Body = "Talent Management Survey"
   .Attachments.Add Pathtofile
   
   For Each objOutlookRecip In .Recipients
       objOutlookRecip.Resolve
   Next
   If DisplayMsg Then
       .Display
   Else
       .Send
   End If
End With
Set objOutlook = Nothing

Application.DisplayAlerts = False
wb.Sheets("Sheet1").Visible = xlSheetVisible
wb.Sheets("Intro").Delete
wb.Sheets("Survey").Delete
wb.Sheets("Model").Delete
wb.Sheets("Summary Report").Delete
wb.Sheets("Results").Delete

ActiveWorkbook.Save


End Sub


