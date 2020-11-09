Imports System.Text.RegularExpressions
Imports Gecko
Imports Newtonsoft.Json.Linq
Public Class XtraForm1
#Region "Check Login"
    Sub Xpath_Method(ByVal Xpath As String, ByVal mth As String, Optional valx As String = "")
        Dim element = WebBrowser1.DomDocument.EvaluateXPath(Xpath)
        Dim nodes = element.GetNodes()
        Dim paginElement = nodes.Select(Function(x) TryCast(x, GeckoElement)).FirstOrDefault()
        If paginElement Is Nothing Then
            MsgBox("NULL")
        Else
            Dim elementHTML = TryCast(paginElement, GeckoHtmlElement)
            If elementHTML IsNot Nothing Then
                If mth = "click" Then
                    elementHTML.Click()
                ElseIf mth = "set" Then
                    elementHTML.SetAttribute("value", valx)

                End If True Then

End If
            End If
        End If
    End Sub

    Sub CheckLogin()
        WebBrowser1.Navigate("https://bulksell.ebay.com/ws/eBayISAPI.dll?SingleList&sellingMode=AddItem")
    End Sub
    Private Sub WebBrowser1_DOMContentLoaded(sender As Object, e As DomEventArgs) Handles WebBrowser1.DOMContentLoaded
        If WebBrowser1.Url.ToString.ToLower.Contains("signin.ebay.com") Then
            Xpath_Method("//*[@id='userid']", "set", "zogakum@mymail90.com")
            Xpath_Method("//*[@id='pass']", "set", "Aa123456789")
            Xpath_Method("//*[@id='sgnBt']", "click")
        ElseIf WebBrowser1.Url.ToString.ToLower.Contains("bulksell.ebay.com") Then
            '
        End If
    End Sub
#End Region

    Private Sub BtnUrlCorrection(sender As Object, e As EventArgs) Handles Btn_UrlCorrection.Click
        Try
            For Each currentLine As String In MemoEdit1.Lines
                If currentLine.Contains("/ref=") Then
                    Dim str1 As String = currentLine
                    Dim str2 As String = ""
                    str2 = str1.Substring(str1.LastIndexOf("/ref="))
                    Dim Final As String = str1.Replace(str2, "")
                    MemoEdit1.Text = MemoEdit1.Text.Replace(currentLine, Final)
                End If

            Next
        Catch ex As Exception
        End Try
    End Sub
    Private Async Sub Btn_Start_ClickAsync(sender As Object, e As EventArgs) Handles Btn_Start.Click
        Dim WBclass As WBBclass
        For Each urls As String In MemoEdit1.Lines
            WBclass = New WBBclass
            WBclass.Navigation1(urls)
            Await Task.Delay(1000)
        Next
    End Sub

End Class