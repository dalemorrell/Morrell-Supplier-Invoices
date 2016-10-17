Imports System.Net
Imports System.IO
Imports System.Threading.Tasks

''' <summary>
''' Specific implementation of BillsFromWebsite for Tradelink
''' </summary>
''' <remarks></remarks>
Public Class Tradelink
    Inherits BillsFromWebsite
#Region "Constructors"
    Sub New(supplier As Supplier)
        MyBase.New(supplier)
    End Sub
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
        End If
        Dim webRequest As HttpWebRequest
        Dim responseReader As StreamReader
        Dim responseData As String
        Dim postData As String
        Dim requestWriter As StreamWriter
        Dim response As WebResponse

        'Get the login page
        webRequest = Net.WebRequest.Create("https://tradedoor.tradelink.com.au")
        response = Await webRequest.GetResponseAsync
        responseReader = New StreamReader(response.GetResponseStream())
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        'have a cookie container for the forms auth cookie
        Dim retVal As CookieContainer
        retVal = New CookieContainer()

        'post the login form
        webRequest = Net.WebRequest.Create("https://tradedoor.tradelink.com.au/users/login")
        webRequest.Method = "POST"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = retVal

        'write the form values into the request message
        requestWriter = New StreamWriter(Await webRequest.GetRequestStreamAsync())
        postData = String.Format("_method=POST&data%5BUser%5D%5Busername%5D={0}&data%5BUser%5D%5Bpassword%5D={1}&data%5BUser%5D%5Bauto_login%5D=0&data%5BUser%5D%5Bauto_login%5D=1&data%5Bgmt_offset_hours%5D=10", Supplier.WebsiteUserID, Supplier.Password, "login")
        Await requestWriter.WriteAsync(postData)
        requestWriter.Close()

        'check the contents for an error
        response = Await webRequest.GetResponseAsync
        responseReader = New StreamReader(response.GetResponseStream())
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()
        If responseData.Contains("The username or password you entered is incorrect.") Then
            Return Nothing
        Else
            Return retVal
        End If
    End Function

    ''' <summary>
    ''' Extract Invoice information from the Html
    ''' </summary>
    ''' <param name="Html"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetBillData(Html As String) As Bill
        Dim retVal As New Bill

        Dim doc As HtmlAgilityPack.HtmlDocument
        doc = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(Html)

        'Get the table containing the address info
        Dim divElement As HtmlAgilityPack.HtmlNode
        divElement = doc.GetElementbyId("onlineinvoice_header")
        Dim cellElements = From td In divElement.Descendants("td")
                           Where td.GetAttributeValue("class", String.Empty) = "invinfo"
                           Select td
        With retVal
            Dim addressline As String
            Try
                'Process the address to fit in the table limits
                addressline = HtmlAgilityPack.HtmlEntity.DeEntitize(cellElements(2).InnerText).Replace("DELIVER TO", "").Trim
                .ShipToAddressLine1 = addressline.Substring(0, Math.Min(256, addressline.Length))
            Catch ex As NullReferenceException
                'sometimes latter AddressLines give us nothing not string.empty
                'when this happens take what we have got so far and move on
            End Try
            Dim tempStr As String()
            tempStr = Split(HtmlAgilityPack.HtmlEntity.DeEntitize(cellElements(1).InnerHtml), "<br>")
            For i = 0 To tempStr.Count - 1
                tempStr(i) = tempStr(i).Trim
            Next
            tempStr(2) = tempStr(2).Substring(0, Math.Min(8, tempStr(2).Length)).ToUpper
            Dim oRE = New Regex("[\W_]")
            tempStr(2) = oRE.Replace(tempStr(2), "")
            .OrderNo = tempStr(2)
            If String.IsNullOrWhiteSpace(.OrderNo) Then
                .OrderNo = "NoOrder"
            End If
            .BillDate = tempStr(1).Replace("-", "/").Insert(tempStr(1).Length - 2, "20")
            .BillNo = tempStr(0).Substring(0, Math.Min(20, tempStr(0).Length))
        End With

        'Get the table containing the line detail
        divElement = doc.GetElementbyId("invoicedetail")
        Dim rowElements = From tr In divElement.Element("table").Elements("tr")
                          Select tr

        Dim Total As String = (divElement.Element("table").Element("tfoot").Elements("tr")(0).Elements("td")(2).InnerText).Trim
        If Total.EndsWith("CR") Then
            retVal.TotalLines = Math.Round(-CSng(Total.Replace("CR", "")), 2)
        Else
            retVal.TotalLines = Math.Round(CSng(Total), 2)
        End If

        Total = (divElement.Element("table").Element("tfoot").Elements("tr")(1).Elements("td")(2).InnerText).Trim
        If Total.EndsWith("CR") Then
            retVal.TotalTax = Math.Round(-CSng(Total.Replace("CR", "")), 2)
        Else
            retVal.TotalLines = Math.Round(CSng(Total), 2)
        End If

        For Each tr In rowElements
            Dim line = New BillLine
            With line
                Dim tempStr As String
                tempStr = HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(0).InnerText).Trim
                If tempStr = "" Then
                    tempStr = "\c"
                End If
                .ItemNo = tempStr.Substring(0, Math.Min(30, tempStr.Length))
                .Name = HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(2).InnerText).Trim
                Try
                    .Quantity = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(5).InnerText).Trim)
                Catch e As InvalidCastException
                    .Quantity = 1
                End Try
                If .Quantity = 0 Then
                    .Quantity = 1
                End If
                Try
                    .TaxExclusiveUnitPrice = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(6).InnerText).Trim)
                Catch e As InvalidCastException
                    .TaxExclusiveUnitPrice = 0
                End Try
                Try
                    .TaxInclusiveUnitPrice = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(4).InnerText).Trim)
                Catch e As InvalidCastException
                    .TaxInclusiveUnitPrice = 0
                End Try
                Try
                    Dim ExTaxTotal As String = HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(7).InnerText).Trim
                    If ExTaxTotal.Contains("CR") Then
                        .TaxExclusiveTotal = -CDbl(ExTaxTotal.Substring(0, Len(ExTaxTotal) - 3))
                    Else
                        .TaxExclusiveTotal = CDbl(ExTaxTotal)
                    End If
                Catch e As InvalidCastException
                    .TaxExclusiveTotal = 0
                End Try
                .TaxInclusiveTotal = .Quantity * .TaxInclusiveUnitPrice
                .Received = .Quantity
            End With
            retVal.BillLines.Add(line)
        Next
        Return retVal
    End Function

    ''' <summary>
    ''' Given an InvoiceAccess (POST or GET) get the Html
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Async Function GetBillHtmlAsync(RequestData As String) As Task(Of String)
        Dim retVal As String
        'send the cookie with the request to invoices
        Dim webRequest As HttpWebRequest
        'goto the invoices page
        webRequest = Net.WebRequest.Create(RequestData)
        webRequest.Method = "GET"
        webRequest.CookieContainer = Await Cookies()

        'Read the response
        Dim responseReader As StreamReader
        Dim response = Await webRequest.GetResponseAsync
        responseReader = New StreamReader(response.GetResponseStream())
        retVal = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        Return retVal
    End Function

    Public Overrides Async Function GetBillAsync(BillNo As String) As Task(Of iBillsFromWebsite)
        Dim html = Await GetBillHtmlAsync(String.Format("https://tradedoor.tradelink.com.au/invoices/view/{0}", BillNo))
        Return GetBillData(html)
    End Function


    ''' <summary>
    ''' Get the Invoice numbers from the website
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Overrides Async Function GetBillNumbersAsync() As Task(Of List(Of String))
        'send the cookie with the request to months invoices
        Dim webRequest As HttpWebRequest
        'goto the invoices page
        Dim url As String
        url = String.Format("https://tradedoor.tradelink.com.au/invoices?unique_download_id_searched=unique_download_id_searched_1339204502802&from_invoice_date={0:dd}%2F{0:MM}%2F{0:yyyy}&to_invoice_date={1:dd}%2F{1:MM}%2F{1:yyyy}&Invoice__credit_note_number=&from_ar_doc_number=&to_ar_doc_number=&open_amount=&Invoice__sales_order_number=&ProjectItem__project_id=&Invoice__customer_order_ref=&Invoice__branch_number=&DeliveryDetail__details=&download_status=2&InvoiceTransaction__product_code=&limit=500",
                            StartDate, EndDate)
        webRequest = Net.WebRequest.Create(url)
        webRequest.Method = "GET"
        webRequest.CookieContainer = Await Cookies()

        'Read the response
        Dim responseReader As StreamReader
        Dim response = Await webRequest.GetResponseAsync()
        responseReader = New StreamReader(response.GetResponseStream())
        Dim responseData As String
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        If responseData.Contains("No invoices were found that matched the search criteria.") Then
            Return New List(Of String)
        End If

        'Convert the response into a HtmlDocument
        Dim doc As HtmlAgilityPack.HtmlDocument
        doc = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(responseData)
        Dim tableRows = doc.DocumentNode.Descendants("table")(1).ChildNodes(1).ChildNodes

        Dim retVal = New Dictionary(Of String, InvoiceAccess)
        For i = 0 To tableRows.Count - 2
            Dim invNo As String = tableRows(i).Descendants("a")(0).InnerText.Trim
            retVal.Add(invNo,
                       New InvoiceAccess With {
                            .Type = InvoiceAccess.AccessType.GET_Request,
                            .RequestData = "https://tradedoor.tradelink.com.au" & tableRows(i).Descendants("a")(0).Attributes("href").Value
                            })
        Next
        Return retVal.Keys.ToList
    End Function
#End Region
End Class
