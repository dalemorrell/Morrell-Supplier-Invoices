
Imports Morrell_Supplier_Invoices
Imports MYOB.AccountRight.SDK.Contracts.Version2

<DataContract>
Public Class Supplier

    <DataMember>
    Public Property CompanyName As String
    <DataMember>
    Public Property FirstName As String
    <DataMember>
    Public Property ItemPrefix As String
    <DataMember>
    Public Property LastName As String
    <IgnoreDataMember>
    Public Property Password As String
    <DataMember>
    Public Property UID As Guid
    <IgnoreDataMember>
    Public Property WebsiteUserID As String
End Class


