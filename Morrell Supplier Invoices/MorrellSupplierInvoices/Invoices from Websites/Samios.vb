Imports System.Net
Imports System.IO
Imports HtmlAgilityPack
Imports System.Security
Imports System.Threading.Tasks

Public Class Samios
    Inherits BillsFromWebsite

#Region "Fields"
    Private viewstate As String
    Private eventvalidation As String
    Private eventtarget As String
    Private eventargument As String
#End Region

#Region "Constructors"

    Sub New(supplier As Supplier)
        MyBase.New(supplier)
    End Sub
#End Region

    Private Function ExtractViewState(s As String, Optional viewStateNameDelimiter As String = "__VIEWSTATE") As String
        Dim valueDelimiter = "value="""

        Dim viewStateNamePosition = s.IndexOf(viewStateNameDelimiter)
        If viewStateNamePosition = -1 Then
            Return String.Empty
        End If
        Dim viewStateValuePosition = s.IndexOf(valueDelimiter, viewStateNamePosition)

        Dim viewStateStartPosition = viewStateValuePosition + valueDelimiter.Length
        Dim viewStateEndPosition = s.IndexOf("""", viewStateStartPosition)
        Return HttpUtility.UrlEncode(s.Substring(viewStateStartPosition, viewStateEndPosition - viewStateStartPosition))
    End Function

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

        'have a cookie container for the forms auth cookie
        Dim retVal As CookieContainer
        retVal = New CookieContainer()

        'Get the login page
        webRequest = HttpWebRequest.Create(String.Format("https://online.samios.net.au/sam/login2.aspx?USERNAME={0}&PASSWORD={1}", Me.UserID, Me.Password))
        webRequest.CookieContainer = retVal
        webRequest.UserAgent = "Internet Explorer"
        Dim response = Await webRequest.GetResponseAsync
        responseReader = New StreamReader(response.GetResponseStream())
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        viewstate = ExtractViewState(responseData)
        eventvalidation = ExtractViewState(responseData, "__EVENTVALIDATION")
        eventtarget = ExtractViewState(responseData, "__EVENTTARGET")
        eventargument = ExtractViewState(responseData, "__EVENTARGUMENT")

        'post the login form
        webRequest = HttpWebRequest.Create(String.Format("https://online.samios.net.au/sam/login2.aspx?USERNAME={0}&PASSWORD={1}", Me.UserID, Me.Password))
        webRequest.AllowAutoRedirect = True
        webRequest.UserAgent = "Internet Explorer"
        webRequest.Method = "POST"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = retVal

        'write the form values into the request message
        requestWriter = New StreamWriter(Await webRequest.GetRequestStreamAsync())
        postData = String.Format("__EVENTTARGET={0}&__EVENTARGUMENT={1}&__VIEWSTATE={2}&__EVENTVALIDATION={3}&ctl00%24ContentPlaceHolder1%24TextBox_Code={4}&ctl00%24ContentPlaceHolder1%24TextBox_Password={5}&ctl00%24ContentPlaceHolder1%24Button_Login={6}",
                                 eventtarget, eventargument, viewstate, eventvalidation, Supplier.WebsiteUserID, Supplier.Password, "++Login++")
        Await requestWriter.WriteAsync(postData)
        requestWriter.Close()

        'check the contents for an error
        response = Await webRequest.GetResponseAsync
        responseReader = New StreamReader(response.GetResponseStream())
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        viewstate = ExtractViewState(responseData)
        eventvalidation = ExtractViewState(responseData, "__EVENTVALIDATION")
        eventtarget = ExtractViewState(responseData, "__EVENTTARGET")
        eventargument = ExtractViewState(responseData, "__EVENTARGUMENT")

        If responseData.Contains("There seems to be a problem with the database server") Or
           responseData.Contains("Login failed") Then
            Return Nothing
        Else
            Return retVal
        End If
    End Function

    Protected Function GetBillData(Html As String) As Bill
        Dim retVal As New Bill

        Dim doc = New HtmlDocument()
        doc.LoadHtml(Html)

        'Get the table containing the address info

        With retVal
            Dim tempStr = HtmlEntity.DeEntitize(doc.GetElementbyId("ctl00_ContentPlaceHolder1_FormView1_Label3").InnerText).Trim.ToUpper
            .BillNo = tempStr.Substring(0, Math.Min(20, tempStr.Length))
            tempStr = HtmlEntity.DeEntitize(doc.GetElementbyId("ctl00_ContentPlaceHolder1_FormView1_Label10").InnerText).Trim
            .BillDate = String.Format("{0:dd/MM/yyy}", Date.Parse(tempStr, Culture))
            tempStr = HtmlEntity.DeEntitize(doc.GetElementbyId("ctl00_ContentPlaceHolder1_FormView1_Label5").InnerText).Trim.ToUpper
            Dim oRE = New Regex("[\W_]")
            tempStr = oRE.Replace(tempStr, "")
            .OrderNo = tempStr.Substring(0, Math.Min(8, tempStr.Length))
        End With

        'Get the invoice totals
        With retVal
            .TotalLines = Math.Round(CSng(HtmlAgilityPack.HtmlEntity.DeEntitize(doc.GetElementbyId("ctl00_ContentPlaceHolder1_FormView2_ExtendedInc").InnerText).Trim), 2)
            .TotalTax = Math.Round(CSng(HtmlAgilityPack.HtmlEntity.DeEntitize(doc.GetElementbyId("ctl00_ContentPlaceHolder1_FormView2_TaxTotal").InnerText).Trim), 2)
        End With

        'Get the table containing the line detail
        Dim rows = doc.GetElementbyId("ctl00_ContentPlaceHolder1_GridView1").Elements("tr")

        For i = 1 To rows.Count - 1
            Dim tr = rows(i)
            Dim line = New BillLine
            With line
                Dim tempStr As String
                tempStr = HtmlEntity.DeEntitize(tr.Elements("td")(0).InnerText).Trim
                If tempStr = "" Then
                    tempStr = "\c"
                End If
                .ItemNo = tempStr.Substring(0, Math.Min(30, tempStr.Length))
                tempStr = HtmlEntity.DeEntitize(tr.Elements("td")(1).InnerText).Trim
                tempStr = tempStr.Substring(0, Math.Min(30, tempStr.Length))
                Try
                    .Quantity = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(2).InnerText).Trim)
                Catch ex As Exception
                    .Quantity = 1
                End Try
                Try
                    .Received = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(3).InnerText).Trim)
                Catch e As InvalidCastException
                    .Received = 0
                End Try
                Try
                    .TaxExclusiveUnitPrice = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(6).InnerText).Trim)
                Catch e As InvalidCastException
                    .TaxExclusiveUnitPrice = 0
                End Try
                Try
                    .TaxExclusiveTotal = CDbl(HtmlAgilityPack.HtmlEntity.DeEntitize(tr.Elements("td")(9).InnerText).Trim)
                Catch e As InvalidCastException
                    .TaxExclusiveTotal = 0
                End Try
            End With
            retVal.BillLines.Add(line)
        Next
        Return retVal
    End Function

    Protected Async Function GetBillHtmlAsync(Access As InvoiceAccess) As Task(Of String)
        Dim webRequest As HttpWebRequest
        Dim responseReader As StreamReader
        Dim responseData As String

        webRequest = Net.WebRequest.Create(Access.RequestData)
        webRequest.AllowAutoRedirect = True
        webRequest.UserAgent = "Internet Explorer"
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = Await Cookies()

        'check the contents for an error
        Dim response = Await webRequest.GetResponseAsync()
        responseReader = New StreamReader(response.GetResponseStream())
        responseData = Await responseReader.ReadToEndAsync()
        responseReader.Close()

        Return responseData
    End Function

    Public Overrides Async Function GetBillAsync(BillNo As String) As Task(Of Bill)

    End Function


    Private Function GetBillNumbersFromHtml(Html As String) As Dictionary(Of String, InvoiceAccess)
        'Convert the response into a HtmlDocument
        Dim doc As HtmlDocument
        doc = New HtmlDocument()
        doc.LoadHtml(Html)
        Dim table = doc.DocumentNode.Descendants("table").Where(Function(t) t.GetAttributeValue("class", String.Empty) = "datagrid").FirstOrDefault()
        Dim retVal = New Dictionary(Of String, InvoiceAccess)
        If table IsNot Nothing Then
            Dim anchors = From r In table.Descendants("tr")
                          Where r.GetAttributeValue("class", String.Empty) = "data" Or r.GetAttributeValue("class", String.Empty) = "dataAltRow"
                          Select r.Descendants("a").First

            For Each anchor In anchors
                retVal.Add(anchor.InnerText.Trim,
                           New InvoiceAccess() With {
                               .Type = InvoiceAccess.AccessType.GET_Request,
                               .RequestData = "https://online.samios.net.au/sam/" & anchor.GetAttributeValue("href", String.Empty)
                           })

            Next
        End If
        Return retVal
    End Function

    Protected Overrides Function GetInvoiceNumbersFromWeb() As Dictionary(Of String, InvoiceAccess)
        Dim webRequest As HttpWebRequest
        Dim responseReader As StreamReader
        Dim responseData As String
        Dim postData As String
        Dim requestWriter As StreamWriter
        Dim retVal = New Dictionary(Of String, InvoiceAccess)


        webRequest = Net.WebRequest.Create(String.Format("https://online.samios.net.au/sam/invoicesearch.aspx?page=range&from={0:ddMMMyyyy}&to={1:ddMMMyyyy}", StartDate, EndDate))
        webRequest.AllowAutoRedirect = True
        webRequest.UserAgent = "Internet Explorer"
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        webRequest.CookieContainer = Cookies()

        'check the contents for an error
        responseReader = New StreamReader(webRequest.GetResponse().GetResponseStream())
        responseData = responseReader.ReadToEnd()
        responseReader.Close()

        'check for no invoices
        If responseData.Contains("Your search returned no results.") Then
            Return New Dictionary(Of String, InvoiceAccess)
        End If


        viewstate = ExtractViewState(responseData)
        eventvalidation = ExtractViewState(responseData, "__EVENTVALIDATION")
        eventtarget = ExtractViewState(responseData, "__EVENTTARGET")
        eventargument = ExtractViewState(responseData, "__EVENTARGUMENT")

        'Convert the response into a HtmlDocument
        Dim doc As HtmlAgilityPack.HtmlDocument
        doc = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(responseData)
        Dim table = doc.DocumentNode.Descendants("table").Where(Function(t) t.GetAttributeValue("class", String.Empty) = "datagrid").FirstOrDefault()
        Dim pagerRow = table.Descendants("tr").Where(Function(t) t.GetAttributeValue("class", String.Empty) = "pager").FirstOrDefault
        Dim maxPages As Integer
        Try
            maxPages = (From txt In pagerRow.Descendants()
                        Where txt.NodeType = HtmlNodeType.Text AndAlso Not String.IsNullOrWhiteSpace(txt.InnerText)
                        Select CType(txt.InnerText, Integer)).Max
        Catch ex As NullReferenceException
            maxPages = 1
        End Try

        Dim pages As New List(Of String)
        pages.Add(responseData)

        For i = 2 To maxPages
            webRequest = HttpWebRequest.Create(String.Format("https://online.samios.net.au/sam/invoicesearch.aspx?page=range&from={0:ddMMMyyyy}&to={1:ddMMMyyyy}",
                                               StartDate, EndDate))
            webRequest.AllowAutoRedirect = True
            webRequest.UserAgent = "Internet Explorer"
            webRequest.Method = "POST"
            webRequest.ContentType = "application/x-www-form-urlencoded"
            webRequest.CookieContainer = Cookies()

            'write the form values into the request message
            requestWriter = New StreamWriter(webRequest.GetRequestStream())
            postData = String.Format("__EVENTTARGET={0}&__EVENTARGUMENT={1}&__VIEWSTATE={2}&__EVENTVALIDATION={3}&ctl00%24SearchDropDownList={4}&ctl00%24SearchText={5}&ctl00_PageMenu_ContextData={6}&ctl00%24ProductSearchCategory={7}&ctl00%24ProductSearchText={8}&ctl00_ProductViewMenu_ContextData={9}&ctl00%24ContentPlaceHolder1%24FromDate%24ctl00={10:dd-MMM-yyyy}&ctl00%24ContentPlaceHolder1%24ToDate%24ctl00={11:dd-MMM-yyyy}",
                                     "ctl00%24ContentPlaceHolder1%24GridView1", "Page%24" & i, viewstate, eventvalidation, "order", String.Empty, String.Empty, "*", String.Empty, String.Empty, StartDate, EndDate)
            requestWriter.Write(postData)
            requestWriter.Close()

            'check the contents for an error
            responseReader = New StreamReader(webRequest.GetResponse().GetResponseStream())
            responseData = responseReader.ReadToEnd()
            responseReader.Close()

            viewstate = ExtractViewState(responseData)
            eventvalidation = ExtractViewState(responseData, "__EVENTVALIDATION")
            eventtarget = ExtractViewState(responseData, "__EVENTTARGET")
            eventargument = ExtractViewState(responseData, "__EVENTARGUMENT")
            pages.Add(responseData)
        Next

        For Each p In pages
            Dim pageInvoices = GetBillNumbersFromHtml(p)
            For Each pi In pageInvoices
                retVal.Add(pi.Key, pi.Value)
            Next
        Next

        Return retVal
    End Function
End Class
