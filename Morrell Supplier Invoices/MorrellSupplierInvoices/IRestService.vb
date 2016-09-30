Imports System.ServiceModel

Public Enum Suppliers
    PlumbersSuppliesCoOp
    Tradelink
    Samios
End Enum

<ServiceContract()>
Public Interface IRestService

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}?from={startDate}&to={endDate}&format=*",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetInvoiceNumbers(supplier As String, startDate As Date, endDate As Date) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}?from={startDate}&to={endDate}&format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetInvoiceNumbersXml(supplier As String, startDate As Date, endDate As Date) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}/{invoiceNo}?format=*",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Json)>
    Function GetInvoice(supplier As String, invoiceNo As String) As List(Of String)

    <OperationContract>
    <WebGet(UriTemplate:="{supplier}/{invoiceNo}?format=xml",
            BodyStyle:=WebMessageBodyStyle.Wrapped,
            ResponseFormat:=WebMessageFormat.Xml)>
    Function GetInvoiceXml(supplier As String, invoiceNo As String) As List(Of String)

End Interface
