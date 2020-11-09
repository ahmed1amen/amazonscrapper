Imports System.Net
Imports System.Text.RegularExpressions
Imports Gecko
Imports Gecko.Events

Public Class WBBclass
    Dim T_Category As String
    Dim Html, titelofItem As String
    Sub ADDRows(ByVal urlC As String)
        If urlC = "" Then Exit Sub

        Dim tClient As WebClient = New WebClient
        Dim client As WebClient = New WebClient
        'client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36")
        'client.Headers.Add("Host", "www.amazon.com")
        'client.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8")
        'Dim html As String
        'html = client.DownloadString(urlC)
        Dim doc As New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(urlC)







        '  Dim Product_Titel As HtmlAgilityPack.HtmlNode = doc.DocumentNode.SelectSingleNode("//*[@id='productTitle']")
        Dim landingImage As HtmlAgilityPack.HtmlNode = doc.DocumentNode.SelectSingleNode("//*[@id='landingImage']")

        '
        Dim price As HtmlAgilityPack.HtmlNode = doc.DocumentNode.SelectSingleNode("//*[@id='priceblock_ourprice']")
        Dim Description As HtmlAgilityPack.HtmlNode = doc.DocumentNode.SelectSingleNode("//*[@id='aplus']/div")
        Dim priceF As String

        Try
            priceF = price.InnerText
        Catch ex As Exception

            Dim priceOfDeal As HtmlAgilityPack.HtmlNode = doc.DocumentNode.SelectSingleNode("//*[@id='priceblock_dealprice']")
            priceF = priceOfDeal.InnerText
        End Try




        '   Dim t_Image As Bitmap = Bitmap.FromStream(New MemoryStream(tClient.DownloadData(landingImage.GetAttributeValue("data-old-hires", ""))))
        Dim t_Action As String = "add"

        Dim t_Title As String = landingImage.GetAttributeValue("alt", "")
        Dim t_Subtitle As String = ""
        Dim t_ConditionID As String = "1000"
        Dim t_PicURL As String = landingImage.GetAttributeValue("data-old-hires", "")
        Dim t_Quantity As String = "OK"
        Dim t_Format As String = "FixedPrice"
        Dim t_StartPrice As String = priceF
        Dim t_BuyItNowPrice As String = priceF
        Dim t_Duration As String = "GTC"
        Dim t_ImmediatePayRequired As String = 0
        Dim t_Location As String = "Marshfield,WI"
        Dim t_GalleryType As String = ""
        Dim t_PayPalAccepted As String = "1"
        Dim t_PayPalEmailAddress As String = "adelkaram21@gmail.com"
        Dim t_Description As String
        Try
            t_Description = Description.InnerHtml
        Catch ex As Exception
            Description = doc.DocumentNode.SelectSingleNode("//*[@id='productDescription']")
            t_Description = Description.InnerHtml
        End Try
        If t_Title.Length > 80 Then
            t_Title = t_Title.Substring(0, 80)
        End If
        XtraForm1.DataGridView1.Rows.Add(t_Action, T_Category, t_Title, t_Subtitle, t_Description, t_ConditionID, t_PicURL, t_Quantity, t_Format, t_StartPrice, t_BuyItNowPrice, t_Duration, t_ImmediatePayRequired, t_Location, t_GalleryType, t_PayPalAccepted, t_PayPalEmailAddress)
    End Sub
    Public WEBB1 As New Gecko.GeckoWebBrowser
    Public WEBB2 As New Gecko.GeckoWebBrowser
    Public Property HttpUtility As Object
    Public Property HttpServerUtility As Object
#Region "WebBrowser 1 Geting Product Info"
    Sub Navigation1(ByVal url As String)

        AddHandler WEBB1.ProgressChanged, AddressOf wb_ProgressChanged
        WEBB1.Navigate(url)
    End Sub
    Public Sub Wb_ProgressChanged(ByVal sender As System.Object, ByVal e As GeckoProgressEventArgs)
        Dim element = WEBB1.DomDocument.EvaluateXPath("//*[@id='landingImage']")
        Dim nodes = element.GetNodes()
        Dim paginElement = nodes.Select(Function(x) TryCast(x, GeckoElement)).FirstOrDefault()
        If paginElement Is Nothing Then
        Else
            Dim elementHTML = TryCast(paginElement, GeckoHtmlElement)
            If elementHTML IsNot Nothing Then
                Html = WEBB1.Document.GetElementsByTagName("html")(0).OuterHtml
                titelofItem = elementHTML.GetAttribute("alt")
                '   Dim encodeed As String = HttpServerUtility.UrlEncode("https://bulkedit.ebay.com/findproduct/cat-reco?q=shose" & titelofItem)
                Navigation2("https://bulkedit.ebay.com/findproduct/cat-reco?q=" & titelofItem)
                WEBB1.Stop()
                WEBB1.Dispose()
                RemoveHandler WEBB1.ProgressChanged, AddressOf Wb_ProgressChanged
                Exit Sub
            End If
        End If
        Exit Sub
    End Sub
#End Region
#Region "WebBrowser 1 Getting catogary IDS"
    Public Sub Navigation2(ByVal url As String)
        WEBB2.Navigate(url)
        AddHandler WEBB2.DocumentCompleted, AddressOf Wb2_DocumentCompleted

    End Sub
    Public Sub Wb2_DocumentCompleted(ByVal sender As System.Object, ByVal e As GeckoDocumentCompletedEventArgs)
        Dim jasonode As String
        jasonode = WEBB2.Document.GetElementsByTagName("pre")(0).InnerHtml
        Try
            Dim Get_CatID As String = Regex.Match(jasonode, Chr(34) & "id" & Chr(34) & ":(.*?)," & Chr(34)).Groups.Item(1).Value
            T_Category = Get_CatID
        Catch ex As Exception
            T_Category = "Error " & ex.Message
        End Try
        ADDRows(Html)
        WEBB2.Stop()
        WEBB2.Dispose()
        RemoveHandler WEBB2.DocumentCompleted, AddressOf Wb2_DocumentCompleted
    End Sub
#End Region



End Class
