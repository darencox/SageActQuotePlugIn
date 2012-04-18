Imports Act.UI
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace ActQuoteNumbering
    <StandardModule> _
    Friend NotInheritable Class globalModule
        ' Methods
        Public Shared Function NumToFormattedString(ByVal number As Integer) As String
            Return Strings.Format(number, "###00000")
        End Function

        Public Shared Sub RaiseUnhandledException(ByVal ex As Exception, ByVal Optional sAdditionalInfo As String = "")
            Dim text As String = String.Concat(New String() { "An unexpected error has occured. ", sAdditionalInfo, ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10), ex.Message, ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10), ex.ToString })
            If Not Information.IsNothing(ex.InnerException) Then
                [text] = String.Concat(New String() { [text], ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10) & "Inner exception: ", ex.InnerException.Message, ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10), ex.ToString })
            End If
            MessageBox.Show([text], "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End Sub


        ' Fields
        Public Shared actApp As ActApplication
        Public Const c_QuoteNumberIndex As Integer = 2
        Public Const c_SQLDatabase As String = "VSDB"
        Public Const c_SQLPassword As String = "F2d60@X07.8cax4a"
        Public Const c_SQLServer As String = "TS01\ACT7"
        Public Const c_SQLUsername As String = "sa"
    End Class
End Namespace

