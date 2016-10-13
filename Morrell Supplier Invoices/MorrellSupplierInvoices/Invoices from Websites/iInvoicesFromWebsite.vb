Imports System.Threading.Tasks

Public Interface iBillsFromWebsite
    Property Culture As Globalization.CultureInfo

    Property StartDate As Date

    Property EndDate As Date

    ReadOnly Property Supplier As Supplier

    Function GetBillNumbersAsync() As Task(Of List(Of String))

    Function GetBillAsync(BillNo As String) As Task(Of Bill)
End Interface
