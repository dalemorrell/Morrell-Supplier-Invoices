' NOTE: You can use the "Rename" command on the context menu to change the class name "RestService" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select RestService.svc or RestService.svc.vb at the Solution Explorer and start debugging.
Imports System.Globalization
Imports System.IO
Imports System.ServiceModel.Activation
Imports System.Threading.Tasks
Imports System.Xml
Imports System.Xml.Serialization

<AspNetCompatibilityRequirements(RequirementsMode:=AspNetCompatibilityRequirementsMode.Allowed)>
<ServiceBehavior(InstanceContextMode:=InstanceContextMode.Single)>
Public Class RestService
    Implements IRestService
    ''' <summary>
    ''' Stores supplier list
    ''' </summary>
    Private _suppliers As List(Of Supplier) = GetAllSuppliers()
    Private _supplieractions As Dictionary(Of Guid, iBillsFromWebsite) = GetSupplierActions()

#Region "Suppliers"
    ''' <summary>
    ''' Retrieves suppliers from xml file
    ''' </summary>
    ''' <returns></returns>
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

    Private Function GetSupplierActions() As Dictionary(Of Guid, iBillsFromWebsite)
        Dim retVal = New Dictionary(Of Guid, iBillsFromWebsite)
        For Each s In _suppliers
            retVal.Add(s.UID, BillsFromWebsite.Factory(s))
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' Returns all suppliers (JSON format)
    ''' </summary>
    ''' <returns></returns>
    Public Function GetAllSuppliersJson() As List(Of Supplier) Implements IRestService.GetAllSuppliersJson
        Return _suppliers
    End Function

    ''' <summary>
    ''' Returns all suppliers (XML format)
    ''' </summary>
    ''' <returns></returns>
    Public Function GetAllSuppliersXml() As List(Of Supplier) Implements IRestService.GetAllSuppliersXml
        Return _suppliers
    End Function

    ''' <summary>
    ''' Returns the supplier identified
    ''' </summary>
    ''' <param name="id">The UID (from AccountRight) of the supplier</param>
    ''' <returns></returns>
    Private Function GetSupplier(id As String) As Supplier
        Try
            Dim uid = New Guid(id)
            Return _suppliers.FirstOrDefault(Function(s) s.UID = uid)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the supplier identified (JSON format)
    ''' </summary>
    ''' <param name="id">The UID (from AccountRight) of the supplier</param>
    ''' <returns></returns>
    Public Function GetSupplierJson(id As String) As Supplier Implements IRestService.GetSupplierJson
        Return GetSupplier(id)
    End Function

    ''' <summary>
    ''' Returns the supplier identified (XML format)
    ''' </summary>
    ''' <param name="id">The UID (from AccountRight) of the supplier</param>
    ''' <returns></returns>
    Public Function GetSupplierXml(id As String) As Supplier Implements IRestService.GetSupplierXml
        Return GetSupplier(id)
    End Function
#End Region

#Region "Invoices"
    Private Async Function GetInvoiceNumbersAsync(supplierId As String, startDate As String, endDate As String) As Task(Of List(Of String))
        Try
            Dim uid = New Guid(supplierId)
            Dim bfw = _supplieractions(uid)

            If Not Date.TryParse(startDate, bfw.Culture, DateTimeStyles.None, bfw.StartDate) Then
                bfw.StartDate = Today.Date.AddMonths(-1)
            End If
            If Not Date.TryParse(endDate, bfw.Culture, DateTimeStyles.None, bfw.EndDate) Then
                bfw.EndDate = Today.Date
            End If

            Return Await bfw.GetBillNumbersAsync
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Async Function GetInvoiceNumbersJson(supplierId As String, startDate As String, endDate As String) As Task(Of List(Of String)) Implements IRestService.GetInvoiceNumbersJson
        Return Await GetInvoiceNumbersAsync(supplierId, startDate, endDate)
    End Function

    Public Async Function GetInvoiceNumbersXml(supplierId As String, startDate As String, endDate As String) As Task(Of List(Of String)) Implements IRestService.GetInvoiceNumbersXml
        Return Await GetInvoiceNumbersAsync(supplierId, startDate, endDate)
    End Function

    Public Function GetInvoice(supplierId As String, invoiceNo As String) As iBillsFromWebsite
        Try
            Dim uid = New Guid(supplierId)
            Dim bfw = _supplieractions(uid)
            Return bfw.GetBillAsync(invoiceNo)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetInvoiceJson(supplierId As String, invoiceNo As String) As iBillsFromWebsite Implements IRestService.GetInvoiceJson
        Return GetInvoice(supplierId, invoiceNo)
    End Function

    Public Function GetInvoiceXml(supplierId As String, invoiceNo As String) As iBillsFromWebsite Implements IRestService.GetInvoiceXml
        Return GetInvoice(supplierId, invoiceNo)
    End Function

#End Region
End Class
