Public Class Bill
    Public Property ShipToAddressLine1 As String
    Public Property ShipToAddressLine2 As String
    Public Property ShipToAddressLine3 As String
    Public Property ShipToAddressLine4 As String
    Public Property OrderNo As String
    Public ReadOnly Property JobNo As String
        Get
            Try
                'remove non alphanumerics
                Dim oRE = New Regex("[\W_]")
                Dim CleanOrderNo = oRE.Replace(OrderNo, "").ToLower

                'from the given row
                If CleanOrderNo Like "[1-9]#####*" Then
                    'Purchase number contains job in positions 1-6
                    Return CleanOrderNo.Substring(0, 6)
                ElseIf CleanOrderNo Like "#[0-9o][a-z][a-z]*" Or
                       CleanOrderNo Like "#[0-9o]" Or
                       CleanOrderNo Like "#[0-9o]a2*" Or
                       CleanOrderNo Like "#[0-9o]6m*" Or
                       CleanOrderNo Like "#[0-9o]2f*" Or
                       CleanOrderNo Like "#[0-9o]5j*" Then
                    'Purchase number contains job in positions 0 & 1
                    Return CleanOrderNo.Substring(0, 2)
                ElseIf CleanOrderNo Like "[a-z][a-z]#[0-9o]" Or
                       CleanOrderNo Like "[a-z][a-z]#[0-9o]###" Or
                       CleanOrderNo Like "a2#[0-9o]###" Or
                       CleanOrderNo Like "6m#[0-9o]###" Or
                       CleanOrderNo Like "2f#[0-9o]###" Or
                       CleanOrderNo Like "5j#[0-9o]###" Then
                    'Purchase number contains job in positions 2 & 3
                    Return CleanOrderNo.Substring(2, 2)
                ElseIf CleanOrderNo Like "[a-z][a-z]#[0-9o][0-9o]" Then
                    'Maintenance Job number
                    Return CleanOrderNo
                Else
                    'Buggered if I know what the job number is
                    Return String.Empty
                End If
            Catch ex As Exception
                Return String.Empty
            End Try
        End Get
    End Property
    Public Property BillDate As Date
    Public Property BillNo As String
    Public Property Comment As String
    Public Property SupplierId As Guid
    Public Property TotalLines As Double
    Public Property TotalTax As Double
    Public Property BillLines As List(Of BillLine)
    Public Property Html As String
End Class

Public Class BillLine
    Public Property ItemNo As String
    Public Property Quantity As Double
    Public Property Received As Double
    Public Property Name As String
    Public Property TaxExclusiveUnitPrice As Double
    Public Property TaxInclusiveUnitPrice As Double
    Public Property TaxExclusiveTotal As Double
    Public Property TaxInclusiveTotal As Double
End Class
