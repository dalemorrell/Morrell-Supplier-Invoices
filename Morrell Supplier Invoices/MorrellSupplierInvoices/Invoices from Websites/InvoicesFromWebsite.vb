Imports System.Threading.Tasks

''' <summary>
''' Base class to get invoices from websites
''' </summary>
''' <remarks></remarks>
Partial Public MustInherit Class BillsFromWebsite
    Implements iBillsFromWebsite

#Region "Classes & Structures"
    ''' <summary>
    ''' Access method to a particular website
    ''' </summary>
    ''' <remarks></remarks>
    Protected Class InvoiceAccess
        Public Enum AccessType
            POST_Request
            GET_Request
        End Enum

        Public Type As AccessType
        Public RequestData As String
    End Class
#End Region

#Region "Fields & Properties"
    ''' <summary>
    ''' A CultureInfo object set to "en-AU", used for parsing dates
    ''' </summary>
    ''' <remarks>All websites used are Australian Based</remarks>
    Public Property Culture As System.Globalization.CultureInfo = New Globalization.CultureInfo("en-AU") Implements iBillsFromWebsite.Culture

    ''' <summary>
    ''' All retreived invoices are dated on or after this date
    ''' </summary>
    ''' <remarks>Changing this requires the Invoices to be downloaded again</remarks>
    Public Property StartDate() As Date = Today().Date.AddMonths(-1) Implements iBillsFromWebsite.StartDate

    ''' <summary>
    ''' All retrieved invoices are dated on or before this date
    ''' </summary>
    ''' <remarks>Changing this requires the Invoices to be downloaded again</remarks>
    Public Property EndDate() As Date = Today().Date Implements iBillsFromWebsite.EndDate

    ''' <summary>
    ''' Supplier associated with this solution
    ''' </summary>
    ''' <returns></returns>
    Public Property Supplier As Supplier Implements iBillsFromWebsite.Supplier
#End Region

#Region "Methods"
    ''' <summary>
    ''' GetInvoiceNumbers gets the invoice numbers and the details of how to access them off the web site
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Async Function GetBillNumbersAsync() As Task(Of List(Of String)) Implements iBillsFromWebsite.GetBillNumbersAsync

    ''' <summary>
    ''' Gets the data for a specifed invoice contained in the Html
    ''' </summary>
    ''' <param name="BillNo">The Supplier's Bill Number</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Async Function GetBillAsync(BillNo As String) As Task(Of Bill) Implements iBillsFromWebsite.GetBillAsync

#End Region

#Region "Constructors"
    Public Sub New(Supplier As Supplier)
        _Supplier = Supplier
    End Sub
#End Region

#Region "Factory Methods"
    ''' <summary>
    ''' Returns the specific concrete InvoicesFromWebsite class required by the supplier
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Factory(supplier As Supplier) As BillsFromWebsite
        Try
            Select Case supplier.UID
                Case New Guid("8c62b7d8-d517-4341-9543-345471eb074b")
                    Return New PlumbersSuppliesCoOp(supplier)
                    'Case New Guid("8573859c-e659-448c-9cde-dc47801fba01")
                    '   Return New Samios(supplier)
                Case New Guid("ba9cce9b-f216-4298-9015-f58c69070e48")
                    Return New Tradelink(supplier)
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region
End Class
