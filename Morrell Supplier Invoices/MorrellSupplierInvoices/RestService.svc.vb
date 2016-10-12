' NOTE: You can use the "Rename" command on the context menu to change the class name "RestService" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select RestService.svc or RestService.svc.vb at the Solution Explorer and start debugging.
Imports System.IO
Imports System.ServiceModel.Activation
Imports System.Xml
Imports System.Xml.Serialization
Imports Morrell_Supplier_Invoices

<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
<ServiceBehavior(InstanceContextMode:=InstanceContextMode.Single)>
Public Class RestService
    Implements IRestService

    Private Function GetAllSuppliers() As List(Of Supplier)
        Dim retval As List(Of Supplier)
        Dim xmlDocument = New XmlDocument()
        Dim path = AppDomain.CurrentDomain.GetData("DataDirectory").ToString() + "\Suppliers.xml"
        xmlDocument.Load(path)
        Dim xmlString = xmlDocument.OuterXml
        Using read = New StringReader(xmlString)
            Dim outType = GetType(List(Of Supplier))
            Dim serializer = New XmlSerializer(outType)
            Using reader = New XmlTextReader(read)
                retval = serializer.Deserialize(reader)
                reader.Close()
            End Using
            read.Close()
        End Using
        Return retval
    End Function

    Public Function GetAllSuppliersJson() As List(Of Supplier) Implements IRestService.GetAllSuppliersJson
        Return GetAllSuppliers()
    End Function

    Public Function GetAllSuppliersXml() As List(Of Supplier) Implements IRestService.GetAllSuppliersXml
        Return GetAllSuppliers()
    End Function

    Public Function GetInvoiceJson(supplier As String, invoiceNo As String) As List(Of String) Implements IRestService.GetInvoiceJson
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceNumbersJson(supplier As String, startDate As Date, endDate As Date) As List(Of String) Implements IRestService.GetInvoiceNumbersJson
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceNumbersXml(supplier As String, startDate As Date, endDate As Date) As List(Of String) Implements IRestService.GetInvoiceNumbersXml
        Throw New NotImplementedException()
    End Function

    Public Function GetInvoiceXml(supplier As String, invoiceNo As String) As List(Of String) Implements IRestService.GetInvoiceXml
        Throw New NotImplementedException()
    End Function

    Public Function GetSupplierJson(id As String) As Supplier Implements IRestService.GetSupplierJson
        Throw New NotImplementedException()
    End Function

    Public Function GetSupplierXml(id As String) As Supplier Implements IRestService.GetSupplierXml
        Throw New NotImplementedException()
    End Function
End Class
