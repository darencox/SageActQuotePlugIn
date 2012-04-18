Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.Windows.Forms
Imports Act.Framework.Opportunities
Imports Act.Framework
Imports Microsoft.VisualBasic.CompilerServices
Imports System.Reflection
Imports System.Runtime.CompilerServices

Namespace ActQuoteNumbering
Friend NotInheritable Class AppMgr

Public Shared Function OpenWordTemplate(ByVal sTemplatePath As String, oOpp As Opportunity, ByVal sQuoteNumber As String) As String

Dim strFileName As String 
strFileName =  dataModule.GetQuotePath & "Quotation " & sQuoteNumber & ".docx"
'MessageBox.Show("strFile : " + strFileName)

Dim oWord As New Word.Application
Dim oDoc As Word.Document


oWord.Visible = True
'oWord.Visible = False
oDoc = oWord.Documents.Add(sTemplatePath)



oDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
AppMgr.PopulateDocument(oDoc.ActiveWindow, oOpp)
oDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
AppMgr.PopulateDocument(oDoc.ActiveWindow, oOpp)
oDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument
AppMgr.PopulateDocument(oDoc.ActiveWindow, oOpp)

oDoc.SaveAs(strFileName)

'oDoc.Close()
oDoc = Nothing

'oWord.Quit()
oWord = Nothing

Return strFileName

End Function

Private Shared Sub PopulateDocument(ByRef oWindow As Window, ByRef oOpp As Opportunity)

Dim flag As Boolean
oWindow.Selection.Start = 0

Do

 flag = False

 oWindow.Selection.End = oWindow.Selection.StoryLength
 Dim fnd As Object = oWindow.Selection.Find

 fnd.Text = "["
 fnd.Forward = True

If (fnd.Execute()) Then
   Dim intStart As Integer = oWindow.Selection.Start
   oWindow.Selection.End = oWindow.Selection.StoryLength
   Dim fnd2 As Object = oWindow.Selection.Find
   'fnd2.ClearFormatting()
   fnd2.Text = "]"
   fnd2.Forward = True
   If (fnd2.Execute()) Then
     Dim intEnd As Integer = oWindow.Selection.End
     flag = True
     AppMgr.ReadAndReplace(oWindow, intStart, intEnd, oOpp)
   End If
End If

Loop While flag


End Sub

Public Shared Sub ReadAndReplace(ByRef oWindow As Window, ByVal iStart As Integer, ByVal iEnd As Integer, ByRef oOpp As Opportunity)
Try
oWindow.Selection.Start = iStart
oWindow.Selection.End = iEnd
                
Dim text As String = oWindow.Selection.Text
text = text.Substring(1, (text.Length - 2))

                Dim strArray As String() = [text].Split(New Char() { "."c })
                Dim str4 As String = strArray(0)
                Dim sColumn As String = strArray(1)
                Dim contactData As String = ""
                Dim str5 As String = str4.ToUpper
                If (str5 = "CONTACT") Then
		    contactData = dataModule.GetContactData(sColumn, oOpp.GetContacts(Nothing).Item(0).ID)
                ElseIf (str5 = "OPPORTUNITY") Then
                    contactData = dataModule.GetOpportunityData(sColumn, oOpp.ID)
                ElseIf (str5 = "ADDRESS") Then
                    contactData = dataModule.GetAddressData(sColumn, oOpp.GetContacts(Nothing).Item(0).ID)
                Else
                    contactData = (text & " unknown")
                End If
                If sColumn.ToUpper.Contains("DATE") Then
                    Try 
                        contactData = Conversions.ToDate(contactData).ToString("dd/MM/yyyy")
                    Catch exception1 As Exception
                        ProjectData.SetProjectError(exception1)
                        ProjectData.ClearProjectError
                    End Try
                End If
                If (contactData.Length > 0) Then
                    oWindow.Selection.Text = contactData
                Else
                    Try 
                        oWindow.Selection.End += 1
                    Catch exception3 As Exception
                        ProjectData.SetProjectError(exception3)
                        Dim exception As Exception = exception3
                        ProjectData.ClearProjectError
                    End Try
                    oWindow.Selection.Text = ""
                End If
            Catch exception4 As Exception
                ProjectData.SetProjectError(exception4)
                Dim ex As Exception = exception4
                globalModule.RaiseUnhandledException(ex, "unable to replace document data")
                ProjectData.ClearProjectError
            End Try

End Sub

End Class
End Namespace