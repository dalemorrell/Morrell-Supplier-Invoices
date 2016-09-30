' NOTE: You can use the "Rename" command on the context menu to change the class name "RestService" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select RestService.svc or RestService.svc.vb at the Solution Explorer and start debugging.
Imports System.ServiceModel.Activation
Imports Morrell_Supplier_Invoices

<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
<ServiceBehavior(InstanceContextMode:=InstanceContextMode.Single)>
Public Class RestService
    Implements IRestService

    Public Function GetInvoice(supplier As String, invoiceNo As String) As List(Of String) Implements IRestService.GetInvoice
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceNumbers(supplier As String, startDate As Date, endDate As Date) As List(Of String) Implements IRestService.GetInvoiceNumbers
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceNumbersXml(supplier As String, startDate As Date, endDate As Date) As List(Of String) Implements IRestService.GetInvoiceNumbersXml
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceXml(supplier As String, invoiceNo As String) As List(Of String) Implements IRestService.GetInvoiceXml
        Throw New NotImplementedException()
    End Function
End Class
