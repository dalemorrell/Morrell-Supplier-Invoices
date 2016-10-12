Imports System.ServiceModel

Public Enum Suppliers
    PlumbersSuppliesCoOp
    Tradelink
    Samios
End Enum

<ServiceContract()>
Public Interface IRestService

#Region "Suppliers"
    <OperationContract>
    <WebGet(UriTemplate:="suppliers",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetAllSuppliersJson() As List(Of Supplier)

    <OperationContract>
    <WebGet(UriTemplate:="suppliers?format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetAllSuppliersXml() As List(Of Supplier)

    <OperationContract>
    <WebGet(UriTemplate:="suppliers/{id}",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetSupplierJson(id As String) As Supplier

    <OperationContract>
    <WebGet(UriTemplate:="suppliers/{id}?format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetSupplierXml(id As String) As Supplier
#End Region

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}?from={startDate}&to={endDate}&format=*",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetInvoiceNumbersJson(supplier As String, startDate As Date, endDate As Date) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}?from={startDate}&to={endDate}&format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetInvoiceNumbersXml(supplier As String, startDate As Date, endDate As Date) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}/{invoiceNo}?format=*",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetInvoiceJson(supplier As String, invoiceNo As String) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}/{invoiceNo}?format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetInvoiceXml(supplier As String, invoiceNo As String) As List(Of String)

End Interface
