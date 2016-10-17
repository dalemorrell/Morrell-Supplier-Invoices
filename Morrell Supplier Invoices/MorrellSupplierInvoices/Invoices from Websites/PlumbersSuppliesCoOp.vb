Imports System.Net
Imports System.IO
Imports System.Threading.Tasks

''' <summary>
''' Specific implementation of InvoicesFromWebsite for Plumbers Supplies Co-Op
''' </summary>
''' <remarks></remarks>
Public Class PlumbersSuppliesCoOp
    Inherits BillsFromWebsite

#Region "Constructors"
    Sub New(supplier As Supplier)
        MyBase.New(supplier)
    End Sub
#End Region

#Region "Fields and Properties"

#End Region

#Region "Methods"
    Dim _cookies As CookieContainer = Nothing
    Protected Async Function Cookies() As Task(Of CookieContainer)
        If _cookies Is Nothing Then
            _cookies = Await GetCookiesAsync()
        End If
        Return _cookies
    End Function

    ''' <summary>
    ''' Login to website and get cookies
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Async Function GetCookiesAsync() As Task(Of CookieContainer)
        If String.IsNullOrWhiteSpace(Supplier.WebsiteUserID) Or String.IsNullOrWhiteSpace(Supplier.Password) Then
            Return Nothing
        Else
            Dim webRequest As HttpWebRequest
            Dim responseReader As StreamReader
            Dim responseData As String
            Dim postData As String
            Dim requestWriter As StreamWriter

            'Get the login page
            webRequest = Net.WebRequest.Create("http://www.pscoop.com.au/v2/pages/home/memberlogin.php")
            Dim response = Await webRequest.GetResponseAsync()
            responseReader = New StreamReader(response.GetResponseStream())
            responseData = Await responseReader.ReadToEndAsync()
            responseReader.Close()

            'have a cookie container for the forms auth cookie
            Dim retVal As CookieContainer
            retVal = New CookieContainer()

            'post the login form
            webRequest = Net.WebRequest.Create("http://www.pscoop.com.au/v2/pages/members/Validate_Login.php")
            webRequest.Method = "POST"
            webRequest.ContentType = "application/x-www-form-urlencoded"
            webRequest.CookieContainer = retVal

            'write the form values into the request message
            requestWriter = New StreamWriter(webRequest.GetRequestStream())
            postData = String.Format("accNo={0}&accPwd={1}&form_name={2}", Supplier.WebsiteUserID, Supplier.Password, "login")
            Await requestWriter.WriteAsync(postData)
            requestWriter.Close()

            'we don't need the contents, just the cookie it issues
            response = Await webRequest.GetResponseAsync()
            responseReader = New StreamReader(response.GetResponseStream())
            responseData = responseReader.ReadToEnd()
            responseReader.Close()
            If responseData.Contains("Your log in details were incorrect!") Then
                Return Nothing
            Else
                Return retVal
            End If
        End If
    End Function

    ''' <summary>
    ''' Extract Invoice information from the Html
    ''' </summary>
    ''' <param name="Html"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetBillData(Html As String) As Bill
        Dim retVal = New Bill
        retVal.Html = Html

        Dim doc As HtmlAgilityPack.HtmlDocument
        doc = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(Html)

        'Get the table containing the address info
        Dim tableHtmlElement As HtmlAgilityPack.HtmlNode
        tableHtmlElement = doc.GetElementbyId("AutoNumber2")
        Dim tableRowHtmlElements = tableHtmlElement.Elements("tr")
        Dim tableCellHtmlElements = tableRowHtmlElements(0).Elements("td")
        Dim paraHtmlElements = tableCellHtmlElements(3).Descendants("p")
        'With invoiceRow
        Dim addressline(4) As String
        Try
            'Process the address to fit in the table limits
            For i = 1 To Math.Min(paraHtmlElements.Count - 1, 4)
                addressline(i) = HtmlAgilityPack.HtmlEntity.DeEntitize(paraHtmlElements(i).InnerText).Trim.Substring(0, Math.Min(255, HtmlAgilityPack.HtmlEntity.DeEntitize(paraHtmlElements(i).InnerText).Trim.Length))
            Next
            retVal.ShipToAddressLine1 = addressline(1)
            retVal.ShipToAddressLine2 = addressline(2)
            retVal.ShipToAddressLine3 = addressline(3)
            retVal.ShipToAddressLine4 = addressline(4)
        Catch ex As NullReferenceException
            'sometimes latter AddressLines give us nothing not string.empty
            'when this happens take what we have got so far and move on
        End Try
        'End With

        'Get the table containing the header info
        tableHtmlElement = doc.GetElementbyId("AutoNumber3")
        tableRowHtmlElements = tableHtmlElement.Elements("tr")
        tableCellHtmlElements = tableRowHtmlElements(1).Elements("td")
        With retVal
            Dim tempStr As String
            tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(3).InnerText).Trim.ToUpper
            'remove non alphanumerics
            Dim oRE = New Regex("[\W_]")
            tempStr = oRE.Replace(tempStr, "")
            .OrderNo = tempStr.Substring(0, Math.Min(8, tempStr.Length))
            .BillDate = Date.Parse(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(0).InnerText).Trim)
            tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText).Trim
            .BillNo = tempStr.Substring(0, Math.Min(20, tempStr.Length))
            tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(4).InnerText).Trim
            .Comment = tempStr.Substring(0, Math.Min(255, tempStr.Length))
            .SupplierId = Supplier.UID
        End With

        'Get the table containing the line detail
        tableHtmlElement = doc.GetElementbyId("AutoNumber4")
        tableRowHtmlElements = tableHtmlElement.Elements("tr")
        For i As Integer = 1 To tableRowHtmlElements.Count - 1
            tableCellHtmlElements = tableRowHtmlElements(i).Elements("td")
            Dim TotalCelltext As String = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText).Trim
            If HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText).Trim = "SUBTOTAL  EXCLUSIVE OF GST" Then
                Try
                    retVal.TotalLines = Math.Round(CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(4).InnerText).Trim), 2)
                Catch e As InvalidCastException
                    retVal.TotalLines = 0
                End Try
            ElseIf HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText).Trim = "GST AMOUNT FOR THIS SUPPLY" Then
                Try
                    retVal.TotalTax = Math.Round(CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(4).InnerText).Trim), 2)
                Catch e As InvalidCastException
                    retVal.TotalTax = 0
                End Try
            ElseIf String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(0).InnerText)) AndAlso
                String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText)) AndAlso
                String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(2).InnerText)) AndAlso
                String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(3).InnerText)) AndAlso
                String.IsNullOrWhiteSpace(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(4).InnerText)) Then
                'blank line - skip
            Else
                'Line item
                Dim line = New BillLine
                With line
                    Dim tempStr As String
                    tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(0).InnerText).Trim
                    If tempStr = "" Then
                        tempStr = "\c"
                    End If
                    .ItemNo = tempStr.Substring(0, Math.Min(30, tempStr.Length))
                    Try
                        .Quantity = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(2).InnerText).Trim)
                    Catch e As InvalidCastException
                        .Quantity = 0
                    End Try
                    tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(1).InnerText).Trim
                    tempStr = tempStr.Substring(0, Math.Min(255, tempStr.Length))
                    .Name = tempStr
                    Try
                        .TaxExclusiveUnitPrice = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(3).InnerText).Trim)
                    Catch e As InvalidCastException
                        .TaxExclusiveUnitPrice = 0
                    End Try
                    Try
                        .TaxExclusiveTotal = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tableCellHtmlElements(4).InnerText).Trim)
                    Catch e As InvalidCastException
                        .TaxExclusiveTotal = 0
                    End Try
                    .TaxInclusiveTotal = Math.Round(.TaxExclusiveTotal * 1.1, 2)
                End With
                retVal.BillLines.Add(line)
            End If
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' Given an InvoiceAccess (POST or GET) get the Html
    ''' </summary>
    ''' <param name="Access"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Async Function GetBillHtmlAsync(Access As InvoiceAccess) As Task(Of String)
        'send the cookie with the request for the invoice
        Dim webRequest As HttpWebRequest
        webRequest = Net.WebRequest.Create("http://www.pscoop.com.au/v2/pages/members/Invoice_Detail.php")
        webRequest.Method = "POST"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = Await Cookies()
        Dim requestWriter As StreamWriter
        requestWriter = New StreamWriter(Await webRequest.GetRequestStreamAsync())
        Await requestWriter.WriteAsync(Access.RequestData)
        requestWriter.Close()

        'Read the response
        Dim response = Await webRequest.GetResponseAsync
        Dim responseReader As StreamReader
        responseReader = New StreamReader(response.GetResponseStream())

        'Convert to a string
        Dim retVal As String
        retVal = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        Return retVal
    End Function

    Public Overrides Async Function GetBillAsync(BillNo As String) As Task(Of iBillsFromWebsite)
        Dim JavaParam As New InvoiceAccess
        JavaParam.Type = InvoiceAccess.AccessType.POST_Request
        JavaParam.RequestData = String.Format("dbDate={0}&tranNo={1}&action={2}&target={3}", "1610", BillNo, "Invoice_Detail.php", "winInvoice")
        Return GetBillData(Await GetBillHtmlAsync(JavaParam))
    End Function

    ''' <summary>
    ''' Extract the Java parameters from the CoOp's href for the invoice
    ''' </summary>
    ''' <param name="JavaCallString">The href for an invoice on the CoOp's site</param>
    ''' <returns>A CoOpJavaParameter containing the parameters of the Java function</returns>
    ''' <remarks></remarks>
    Private Function GetJavaParameters(JavaCallString As String) As InvoiceAccess
        Dim h() As String
        h = JavaCallString.Split(""",""")
        Dim JavaParam As New InvoiceAccess
        With JavaParam
            .Type = InvoiceAccess.AccessType.POST_Request
            .RequestData = String.Format("dbDate={0}&tranNo={1}&action={2}&target={3}", h(1), h(3), h(5), h(7))
        End With
        Return JavaParam
    End Function

    ''' <summary>
    ''' Gets the Invoice number from the JavaCall
    ''' </summary>
    ''' <param name="JavaCallString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInvoiceNumber(JavaCallString As String) As String
        Dim h() As String
        h = JavaCallString.Split(""",""")
        Return h(3)
    End Function

    ''' <summary>
    ''' GetInvoiceNumbers scrapes the Java parameters for the given month
    ''' </summary>
    ''' <param name="month">A DateTime to determine the month to be queried from the CoOp's website</param>
    ''' <returns>A List containing the Java parameters of all invoices in the month</returns>
    ''' <remarks></remarks>
    Private Overloads Async Function GetBillNumbersAsync(month As Date) As Task(Of Dictionary(Of String, InvoiceAccess))
        'send the cookie with the request to months invoices
        Dim webRequest As HttpWebRequest
        webRequest = Net.WebRequest.Create("http://www.pscoop.com.au/v2/pages/members/Invoice_List.php")
        webRequest.Method = "POST"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = Await Cookies()
        Dim monthstring As String
        monthstring = WebUtility.HtmlEncode(String.Format("{0:MM/yyyy}", month))
        Dim postData As String
        postData = String.Format("form_name={0}&enquiryType={1}&enquiryInput={2}", "invoice_enquiry", "byMonth", monthstring)
        Dim requestWriter As StreamWriter
        requestWriter = New StreamWriter(Await webRequest.GetRequestStreamAsync())
        Await requestWriter.WriteAsync(postData)
        requestWriter.Close()

        'Read the response
        Dim response = Await webRequest.GetResponseAsync
        Dim responseReader As StreamReader
        responseReader = New StreamReader(response.GetResponseStream())
        Dim responseData As String
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        'Convert the response into a HtmlDocument
        Dim doc As HtmlAgilityPack.HtmlDocument
        doc = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(responseData)

        Dim allinvoices As IEnumerable(Of HtmlAgilityPack.HtmlNode)
        allinvoices = From a In doc.DocumentNode.Descendants
                      Where (a.Name Like "[aA]" AndAlso
                      a.Attributes("href").Value Like "Javascript:SubmitForm(*Invoice_Detail.php*winInvoice*)")
                      Select a

        'Query the document for the invoices.
        Dim invoiceNumbers As Dictionary(Of String, InvoiceAccess)
        invoiceNumbers = New Dictionary(Of String, InvoiceAccess)
        For Each n In allinvoices
            Dim InvParam As InvoiceAccess
            InvParam = GetJavaParameters(n.Attributes("href").Value)
            Dim InvNo As String
            InvNo = GetInvoiceNumber(n.Attributes("href").Value)
            Dim invoiceDate As Date
            Try
                invoiceDate = Date.Parse(n.ParentNode.ParentNode.Descendants("td").FirstOrDefault.InnerText, Culture)
                If invoiceDate >= StartDate And invoiceDate <= EndDate Then
                    invoiceNumbers.Add(InvNo, InvParam)
                End If
            Catch ex As ArgumentNullException
            Catch ex As FormatException
            End Try
        Next
        Return invoiceNumbers
    End Function

    ''' <summary>
    ''' Get the Invoice numbers from the website
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Overrides Async Function GetBillNumbersAsync() As Task(Of List(Of String))
        Dim retVal As New List(Of String)
        Dim sd As Date
        sd = Date.Parse(String.Format("1/{0:MM/yyyy}", StartDate))
        Do While sd <= EndDate
            Dim thisMonth = Await GetBillNumbersAsync(sd)
            retVal.AddRange(thisMonth.Keys)
            sd = sd.AddMonths(1)
        Loop
        Return retVal
    End Function
#End Region
End Class