Imports Act.Framework.Contacts
Imports Act.Framework.Histories
Imports Act.Framework.Opportunities
Imports Act.Framework.SupplementalFiles
Imports Act.UI
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Interop.IWshRuntimeLibrary

Namespace ActQuoteNumbering
    Public Class actClass
        Implements IPlugin
        ' Methods
        Public Sub actApp_AfterLogon(ByVal sender As Object, ByVal e As EventArgs)
            If (globalModule.actApp.ActFramework.CurrentDatabase = "Colporteur_server") Then  'For dev schema name is ActDb. For CP it is Colporteur_server
                AddHandler globalModule.actApp.ActFramework.Opportunities.AddOpportunity, New OpportunityManager.AddOpportunityEvent(AddressOf Me.actapp_OppAdded)
                Dim lQuoteNo As Long = 0
                Dim sQuotePath As String = ""
                Dim sTemplPath As String = ""
                If Not dataModule.LoadConfig(lQuoteNo, sQuotePath, sTemplPath) Then
                    MessageBox.Show("The quote number system is not yet configured.  Please select Quote Number Config from the tools menu to configure the system before entering any new opportunities.", "", MessageBoxButtons.OK)
                End If
            End If
        End Sub

        Public Sub actapp_OppAdded(ByVal opportunityName As String, ByVal opportunityID As Guid)
            Dim lQuoteNo As Long = 0
            Dim sQuotePath As String = ""
            Dim sTemplPath As String = ""
            If dataModule.LoadConfig(lQuoteNo, sQuotePath, sTemplPath) Then
                Me.CreateNewQuote(opportunityID)
            Else
                MessageBox.Show("The quote number system is not yet configured.  This opportunity will not be automatically numbered.", "", MessageBoxButtons.OK)
            End If
        End Sub

        Public Sub AddShortcutToContact(ByVal sQuoteNo As String, ByVal sShortcut As String, ByRef oContact As Contact)
            Try 
                Dim regarding As String = sQuoteNo
                Dim display As String = ("Quote " & sQuoteNo)
                Dim att As Attachment = globalModule.actApp.ActFramework.SupplementalFileManager.CreateAttachment(AttachmentMate.History, sShortcut, display, False)
                globalModule.actApp.ActFramework.Histories.CreateHistory(oContact, Guid.Empty, New HistoryType(SystemHistoryType.Library), False, DateTime.Now, DateTime.Now, regarding, String.Empty, att).Update
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to add quote to contact")
                ProjectData.ClearProjectError
            End Try
        End Sub

        private Sub CreateNewQuote(ByVal opportunityID As Guid)
            Try 
		Dim s As String = opportunityID.ToString()
                Dim oOpp As Opportunity = Me.GetOpportunity(opportunityID)

		If (oOpp.GetContacts(Nothing).Count <>0) Then
                   Dim nextQuoteNumber As String = Me.GetNextQuoteNumber
   	           oOpp.Fields.Item("TBL_OPPORTUNITY.USER3", True) = nextQuoteNumber
                   Dim opportunity2 As Opportunity = oOpp
                   opportunity2.Name = (opportunity2.Name & " (" & nextQuoteNumber & ")")
		   opportunity2.Update()
                   Dim sFullObjectPath As String = AppMgr.OpenWordTemplate(Me.GetUserTemplate, oOpp, nextQuoteNumber)
		   'MessageBox.Show("Document : " + sFullObjectPath)
                   Dim sShortcut As String = Me.CreateShortcut(sFullObjectPath, dataModule.GetQuotePath, ("Quote" & nextQuoteNumber))
		   Dim oContact As Contact = oOpp.GetContacts(Nothing).Item(0)
		   Me.AddShortcutToContact(nextQuoteNumber, sShortcut, oContact)
		   System.IO.File.Delete(sShortcut)
		Else
		   MessageBox.Show("No quote created as no contact available")
		End If
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to continue")
                ProjectData.ClearProjectError
            End Try
        End Sub

	Private Function CreateShortcut(ByVal sFullObjectPath As String, ByVal sWorkDir As String, ByVal sLinkName As String) As String
        'Private Function CreateShortCutOnDesktop(ByVal userID As String, ByVal passWord As String) As Boolean
	Dim strReturn AS String
	strReturn = ""
	  Try

	 Dim WshShell As WshShellClass = New WshShellClass
         'Dim DesktopDir As String = CType(WshShell.SpecialFolders.Item("Desktop"), String)
         Dim shortCut As Interop.IWshRuntimeLibrary.IWshShortcut
	 Dim linkPath As String 
	 linkPath = Application.CommonAppDataPath & "\" & sLinkName & ".lnk"

         ' short cut files have a .lnk extension
	 
         'shortCut = CType(WshShell.CreateShortcut(DesktopDir & "\MyNewShortcut.lnk"), IWshRuntimeLibrary.IWshShortcut)
	 shortCut = CType(WshShell.CreateShortcut(linkPath), Interop.IWshRuntimeLibrary.IWshShortcut)

         ' set the shortcut properties
         With shortCut
            .TargetPath = sFullObjectPath
            .WindowStyle = 1
            .Description = "ACT! Quote!"
            .WorkingDirectory = sWorkDir
            .Save() ' save the shortcut file
         End With
	 strReturn = linkPath
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to create shortcut")
                ProjectData.ClearProjectError
      End Try
      Return strReturn
   End Function

        Private Function GetNextQuoteNumber() As String
            Dim str As String = ""
            Try 
                str = globalModule.NumToFormattedString(CInt(dataModule.GetNewNumberFromDB))
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to acquire quote number")
                ProjectData.ClearProjectError
            End Try
            Return str
        End Function

        Private Function GetOpportunity(ByVal opportunityID As Guid) As Opportunity
            Dim opportunities As OpportunityList
            Dim oppIDs As Guid() = New Guid(1) {}
            Dim opportunity As Opportunity = Nothing
            Try 
                oppIDs(0) = opportunityID
                opportunities = globalModule.actApp.ActFramework.Opportunities.GetOpportunities(oppIDs, Nothing)
		'MessageBox.Show("Opportunity count : " & opportunities.Count)
                If (opportunities.Count > 0) Then

                    Return opportunities.Item(0)
                End If
                opportunity = Nothing
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to acquire opportunity")
                ProjectData.ClearProjectError
            Finally
                oppIDs = Nothing
                opportunities = Nothing
            End Try
            Return opportunity
        End Function

        Private Function GetUserTemplate() As String
            Dim str As String = ""
            Try 
                Dim displayName As String = globalModule.actApp.ActFramework.CurrentUser.DisplayName
                Dim lastTemplate As String = dataModule.GetLastTemplate(displayName)
                Dim path As String = ""
                If (System.IO.File.Exists(lastTemplate) AndAlso (MessageBox.Show(("Last template used was: " & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10) & lastTemplate & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10) & "Use this template?"), "ACT!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)) Then
                    path = lastTemplate
                End If
                If ((path.Length = 0) Or Not System.IO.File.Exists(path)) Then
                    Dim str5 As String = (dataModule.GetTemplatePath & displayName)
                    If Not Directory.Exists(str5) Then
                        Directory.CreateDirectory(str5)
                    End If
                    Dim dialog As New OpenFileDialog
                    Try 
                        Dim dialog2 As OpenFileDialog = dialog
                        dialog2.CheckFileExists = True
                        dialog2.Multiselect = False
                        dialog2.InitialDirectory = str5
                        dialog2.Filter = "Template files (*.dotx)|*.dotx|All files|*.*"
                        Do
                            If (dialog2.ShowDialog = DialogResult.OK) Then
                                path = dialog2.FileName
                            End If
                        Loop While Not System.IO.File.Exists(path)
                        dialog2 = Nothing
                    Catch exception1 As Exception
                        ProjectData.SetProjectError(exception1)
                        Dim ex As Exception = exception1
                        globalModule.RaiseUnhandledException(ex, "")
                        ProjectData.ClearProjectError
                    Finally
                        dialog = Nothing
                    End Try
                End If
                dataModule.StoreLastTemplate(displayName, path)
                str = path
            Catch exception3 As Exception
                ProjectData.SetProjectError(exception3)
                Dim exception2 As Exception = exception3
                globalModule.RaiseUnhandledException(exception2, "unable to get template")
                ProjectData.ClearProjectError
            End Try
            Return str
        End Function

        Public Sub OnLoad(ByVal oApp As ActApplication) Implements IPlugin.OnLoad
            globalModule.actApp = oApp
            AddHandler globalModule.actApp.AfterLogon, New EventHandler(AddressOf Me.actApp_AfterLogon)
        End Sub

        Public Sub OnUnLoad() Implements IPlugin.OnUnLoad
        End Sub

        Public Sub quoteNumberConfig(ByVal [text] As String)
        End Sub

    End Class
End Namespace

