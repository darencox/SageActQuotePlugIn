Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports Act.UI

Namespace ActQuoteNumbering
    <StandardModule> _
    Friend NotInheritable Class dataModule
        ' Methods
        Private Shared Function ConnectToDB() As SqlConnection
            Dim connection As SqlConnection = Nothing
            'Dim connectionString As String = "server=TS01\ACT7;uid=sa;pwd=F2d60@X07.8cax4a;database=VSDB"
            'Dim connectionString As String = "server=127.0.0.1\ACT7;User ID=Admin;Password=Admin;Database=ActDb"
	    'Dim connectionString As String =  "server=.\ACT7;uid=Admin;pwd=Admin;database=ActDb"
	    'Dim connectionString As String =  "server=.\ACT7;Integrated Security=SSPI;database=ActDb" 'use for development
	    'Dim connectionString As String  = "Data Source=.\ACT7;Initial Catalog=ActDb;User Id=Admin;Password=Admin;Connect Timeout=30"
	    'Dim connectionString As String  = "server=Local\SQLEXPRESS;Password=Admin;User ID=Admin;Initial Catalog=ActDb"
	    'Dim connectionString As String =  "server=ELSOXFSQLD005\DEV02;Integrated Security=SSPI;Initial Catalog=AfpReports" 'This one works
	    'Dim  connectionString As String = "server=DELL-SERVER02\ACT7;Integrated Security=SSPI;database=colporteur_server" 'Testing on Colporteur setup
	    Dim connectionString As String = "server=DELL-SERVER02\ACT7;uid=cp;pwd=colporteur123;database=Colporteur_server"   'Testing on Colporteur setup. Created cp acocunt on Ms Sql
            Try 
                connection = New SqlConnection(connectionString)
                connection.Open
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim ex As Exception = exception1
                globalModule.RaiseUnhandledException(ex, "unable to connect to database")
                ProjectData.ClearProjectError
            End Try
            Return connection
        End Function

        Public Shared Function GetAddressData(ByVal sColumn As String, ByVal contactID As Guid) As String
            Dim str As String = ""
            Dim cmdText As String = (("SELECT " & sColumn & " FROM TBL_ADDRESS") & " WHERE CONTACTID = '" & contactID.ToString & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        str = reader.Item(0).ToString
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "unable to get address data")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return str
        End Function

        Public Shared Function GetContactData(ByVal sColumn As String, ByVal contactID As Guid) As String
            Dim str As String = ""
            Dim cmdText As String = (("SELECT " & sColumn & " FROM TBL_CONTACT") & " WHERE CONTACTID = '" & contactID.ToString & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        str = reader.Item(0).ToString
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "unable to get contact data")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return str
        End Function

        Public Shared Function GetLastTemplate(ByVal sUserName As String) As String
            Dim str As String = ""
            Dim userIDFromName As Guid = dataModule.GetUserIDFromName(sUserName)
            Dim cmdText As String = "SELECT LASTTEMPLATE FROM ACS_ACCESSORLASTTEMPLATE"
            cmdText = (cmdText & " WHERE ACCESSORID = '" & userIDFromName.ToString & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        str = reader.Item(0).ToString
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error acquiring last template")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return str
        End Function

        Public Shared Function GetNewNumberFromDB() As Long
            Dim num As Long
            Dim cmdText As String = "SELECT NEXTQUOTENO FROM ACS_CONFIG"
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim command As New SqlCommand(cmdText, expression)
                    Dim reader As SqlDataReader = command.ExecuteReader
                    If reader.Read Then
                        Dim num2 As Long = Conversions.ToLong(reader.Item(0).ToString)
                        reader.Close
                        num = num2
                        num2 = (num2 + 1)
                        Dim cmdUpdate As New SqlCommand
			cmdUpdate.Connection = expression
                        cmdUpdate.CommandText = "UPDATE ACS_CONFIG SET NEXTQUOTENO = " & Conversions.ToString(num2)
                        cmdUpdate.ExecuteNonQuery()
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "unable to get new quote number")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return num
        End Function

        Public Shared Function GetOpportunityData(ByVal sColumn As String, ByVal oppID As Guid) As String
            Dim str As String = ""
            Dim cmdText As String = (("SELECT " & sColumn & " FROM TBL_OPPORTUNITY") & " WHERE OPPORTUNITYID = '" & oppID.ToString & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        str = reader.Item(0).ToString
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "unable to get opportunity data")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return str
        End Function

        Public Shared Function GetQuotePath() As String
            Dim str As String = ""
            Dim cmdText As String = "SELECT QUOTEPATH FROM ACS_CONFIG"
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    reader.Read
                    str = reader.Item(0).ToString
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error acquiring quote path")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            If ((str.Length > 0) AndAlso Not str.EndsWith("\")) Then
                str = (str & "\")
            End If
            Return str
        End Function

        Public Shared Function GetTemplatePath() As String
            Dim str As String = ""
            Dim cmdText As String = "SELECT TEMPLATEPATH FROM ACS_CONFIG"
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    reader.Read
                    str = reader.Item(0).ToString
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error acquiring template path")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            If ((str.Length > 0) AndAlso Not str.EndsWith("\")) Then
                str = (str & "\")
            End If
            Return str
        End Function

        Private Shared Function GetUserIDFromName(ByVal sUserName As String) As Guid
            Dim guid As Guid
            Dim cmdText As String = "SELECT ACCESSORID FROM TBL_ACCESSOR"
            cmdText = (cmdText & " WHERE NAME = '" & sUserName & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        guid = DirectCast(reader.Item(0), Guid)
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error locating user ID")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return guid
        End Function

        Public Shared Function LoadConfig(ByRef Optional lQuoteNo As Long = 0, ByRef Optional sQuotePath As String = "", ByRef Optional sTemplPath As String = "") As Boolean
            Dim flag As Boolean = False
            Dim cmdText As String = "SELECT NEXTQUOTENO, QUOTEPATH, TEMPLATEPATH FROM ACS_CONFIG"
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim reader As SqlDataReader = New SqlCommand(cmdText, expression).ExecuteReader
                    If reader.Read Then
                        lQuoteNo = Conversions.ToLong(reader.Item(0).ToString)
                        sQuotePath = reader.Item(1).ToString
                        sTemplPath = reader.Item(2).ToString
                        If (((lQuoteNo = 0) Or (sQuotePath.Length = 0)) Or (sTemplPath.Length = 0)) Then
                            flag = False
                        Else
                            flag = True
                        End If
                    Else
                        flag = False
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error acquiring config")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return flag
        End Function

        Public Shared Function SaveConfig(ByVal lQuoteNo As Long, ByVal sQuotePath As String, ByVal sTemplPath As String) As Boolean
            Dim flag As Boolean = False
            Dim cmdText As String = "UPDATE ACS_CONFIG SET "
            cmdText = (((cmdText & "NEXTQUOTENO=" & Conversions.ToString(lQuoteNo) & ", ") & "QUOTEPATH='" & sQuotePath & "', ") & "TEMPLATEPATH='" & sTemplPath & "'")
            Dim expression As SqlConnection = dataModule.ConnectToDB
            If Not Information.IsNothing(expression) Then
                Try 
                    Dim command As New SqlCommand(cmdText, expression)
                    Dim num As Integer = command.ExecuteNonQuery
                    If (num = 0) Then
                        cmdText = "INSERT INTO ACS_CONFIG (NEXTQUOTENO, QUOTEPATH, TEMPLATEPATH) "
                        num = New SqlCommand(String.Concat(New String() { cmdText, "VALUES (", Conversions.ToString(lQuoteNo), ", '", sQuotePath, "', '", sTemplPath, "')" }), expression).ExecuteNonQuery
                    End If
                    If (num > 0) Then
                        flag = True
                    End If
                    expression.Close
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim ex As Exception = exception1
                    globalModule.RaiseUnhandledException(ex, "error storing config")
                    ProjectData.ClearProjectError
                Finally
                    expression = Nothing
                End Try
            End If
            Return flag
        End Function

        Public Shared Sub StoreLastTemplate(ByVal sUserName As String, ByVal sTemplate As String)
            Dim userIDFromName As Guid = dataModule.GetUserIDFromName(sUserName)
            If (dataModule.GetLastTemplate(sUserName).Length > 0) Then
                Dim cmdText As String = "UPDATE ACS_ACCESSORLASTTEMPLATE SET "
                cmdText = ((cmdText & "LASTTEMPLATE='" & sTemplate & "' ") & " WHERE ACCESSORID = '" & userIDFromName.ToString & "'")
                Dim expression As SqlConnection = dataModule.ConnectToDB
                If Not Information.IsNothing(expression) Then
                    Try 
                        Dim cmdUpdate As New SqlCommand
			cmdUpdate.Connection = expression
                        cmdUpdate.CommandText = cmdText
                        cmdUpdate.ExecuteNonQuery()
                        expression.Close
                    Catch exception1 As Exception
                        ProjectData.SetProjectError(exception1)
                        Dim ex As Exception = exception1
                        globalModule.RaiseUnhandledException(ex, "error storing last-used template")
                        ProjectData.ClearProjectError
                    Finally
                        expression = Nothing
                    End Try
                End If
            Else
                Dim str2 As String = "INSERT INTO ACS_ACCESSORLASTTEMPLATE (ACCESSORID, LASTTEMPLATE) "
                str2 = String.Concat(New String() { str2, "VALUES ('", userIDFromName.ToString, "', '", sTemplate, "')" })
                Dim connection2 As SqlConnection = dataModule.ConnectToDB
                If Not Information.IsNothing(connection2) Then
                    Try 
                        Dim cmdInsert As New SqlCommand
			cmdInsert.Connection = connection2
                        cmdInsert.CommandText = str2
                        cmdInsert.ExecuteNonQuery()
                        connection2.Close
                    Catch exception3 As Exception
                        ProjectData.SetProjectError(exception3)
                        Dim exception2 As Exception = exception3
                        globalModule.RaiseUnhandledException(exception2, "error storing last-used template")
                        ProjectData.ClearProjectError
                    Finally
                        connection2 = Nothing
                    End Try
                End If
            End If
        End Sub


        ' Fields
        Private Const c_ASR_ID As String = "ACCESSORID"
        Private Const c_ASR_Name As String = "NAME"
        Private Const c_CFG_NextQuoteNo As String = "NEXTQUOTENO"
        Private Const c_CFG_QuotePath As String = "QUOTEPATH"
        Private Const c_CFG_TemplatePath As String = "TEMPLATEPATH"
        Private Const c_ContactID As String = "CONTACTID"
        Private Const c_OpportunityID As String = "OPPORTUNITYID"
        Private Const c_TBL_Accessor As String = "TBL_ACCESSOR"
        Private Const c_TBL_Address As String = "TBL_ADDRESS"
        Private Const c_TBL_Config As String = "ACS_CONFIG"
        Private Const c_TBL_Contact As String = "TBL_CONTACT"
        Private Const c_TBL_Opportunity As String = "TBL_OPPORTUNITY"
        Private Const c_TBL_UserTplt As String = "ACS_ACCESSORLASTTEMPLATE"
        Private Const c_UTP_AccessorID As String = "ACCESSORID"
        Private Const c_UTP_LastTemplate As String = "LASTTEMPLATE"
    End Class
End Namespace

