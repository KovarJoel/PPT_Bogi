Option Explicit On

Imports System.Data.SqlTypes
Imports System.IO
Imports System.Linq.Expressions
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Scripting

Partial Class MyFunctions

    Public Class Mission

        Structure VideoDate
            Dim Year As Integer
            Dim Quarter As Integer
            Dim Number As Integer
        End Structure

        Private Shared Function GetNextSa() As Date

            GetNextSa = Date.Today.AddDays(DayOfWeek.Saturday - Date.Today.DayOfWeek)

        End Function

        Public Shared Function GetVideoDate() As VideoDate

            Dim nextSa As Date = GetNextSa()
            Dim firstMonth As Integer

            With GetVideoDate

                .Year = nextSa.Year

                Select Case nextSa.Month
                    Case 1 To 3
                        .Quarter = 1
                        firstMonth = 1
                    Case 4 To 6
                        .Quarter = 2
                        firstMonth = 4
                    Case 7 To 9
                        .Quarter = 3
                        firstMonth = 7
                    Case 10 To 12
                        .Quarter = 4
                        firstMonth = 10
                    Case Else
                        .Quarter = 0
                        firstMonth = 0
                End Select

            End With

            Dim day = New Date(nextSa.Year, firstMonth, 1)
            Dim count As Integer = 0

            While Not (day.Date = nextSa.Date)
                If day.DayOfWeek = DayOfWeek.Saturday Then
                    count += 1
                End If

                day = day.AddDays(1)

            End While

            GetVideoDate.Number = count + 1

        End Function

        Public Shared Function GetLink() As String

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim videoDate = GetVideoDate()
            Dim url As String = "https://cloud.eud.adventist.org/index.php/s/i9nTwt55bHEmpLJ?path=%2F" & videoDate.Year & "_" & videoDate.Quarter & ".%20Quartal"
            Dim scriptText As String
            Dim fs = IO.File.OpenRead("C:\Users\joelm\source\repos\PPT Bogi\PPT Bogi\MyFunctions\getHTML.js")
            scriptText = New StreamReader(fs).ReadToEnd()

            Dim wb As New WebBrowser()
            wb.ScriptErrorsSuppressed = True
            wb.Navigate(url)

            While wb.ReadyState <> WebBrowserReadyState.Complete
                Application.DoEvents()
            End While

            Dim result As Object = wb.Document.InvokeScript("eval", New Object() {scriptText})
            Dim page As String = ""

            If result IsNot Nothing Then
                page = result.ToString()
            End If

            Dim fso As New FileSystemObject
            Dim file As TextStream
            Files.CreateTextFile(tempDataPath & "site.txt")
            file = fso.OpenTextFile(tempDataPath & "site.txt", IOMode.ForWriting)
            file.Write(page)
            file.Close()

            Exit Function

            'Dim browser As New WebBrowser
            'browser.Navigate(url)
            '
            'While browser.ReadyState <> WebBrowserReadyState.Complete
            '    Application.DoEvents()
            'End While
            '
            'Dim doc As HtmlDocument = browser.Document
            'Dim head As HtmlElement = doc.GetElementsByTagName("head").Item(0)
            'Dim s As HtmlElement = doc.CreateElement("script")
            's.SetAttribute("text", scriptText)
            'head.AppendChild(s)
            '
            'With browser.Document
            '    .GetElementsByTagName("head").Item(0).AppendChild(s)
            'End With
            'Dim result = browser.Document.InvokeScript("getHTML")
            '
            'Exit Function

            'Dim videoDate = GetVideoDate()
            'Dim filePath As String = tempDataPath & "site.html"
            'Dim folderLink As String = "https://cloud.eud.adventist.org/index.php/s/i9nTwt55bHEmpLJ?path=%2F"
            'folderLink &= videoDate.Year & "_" & videoDate.Quarter & ".%20Quartal"
            '
            'Dim request As WebRequest = WebRequest.Create(folderLink)
            'Dim response As WebResponse = request.GetResponse()
            'Dim responseString As String = New StreamReader(response.GetResponseStream()).ReadToEnd()
            '
            'Dim web As New WebBrowser
            'web.DocumentText = responseString


            'Exit Function

            'Dim fso = New FileSystemObject
            'Dim file As TextStream
            'Dim videoDate = GetVideoDate()
            '
            'Dim folderPath As String = tempDataPath & "site.html"
            'Dim folderLink As String = "https://cloud.eud.adventist.org/index.php/s/i9nTwt55bHEmpLJ?path=%2F"
            'folderLink &= videoDate.Year & "_" & videoDate.Quarter & ".%20Quartal"
            '
            'Files.DownloadFileFromLink(folderLink, folderPath)
            '
            'file = fso.OpenTextFile(folderPath)
            '
            '
            '
            'GetLink = "https://cloud.eud.adventist.org/index.php/s/i9nTwt55bHEmpLJ/download?path=%2F"
            'GetLink &= videoDate.Year & "_" & videoDate.Quarter & ".%20Quartal"
            'GetLink &= "&files="


        End Function

        Public Shared Async Function LoadAndSearch(url As String) As Task(Of String)

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            ' Create an HttpClient instance
            Dim httpClient As New Http.HttpClient()

            ' Make an HTTP GET request to the specified URL
            Dim httpResponse As Http.HttpResponseMessage = Await httpClient.GetAsync(url)

            ' Read the response content as a string
            Dim html As String = Await httpResponse.Content.ReadAsStringAsync()

            ' Search for the specified substring in the full HTML content of the website
            Dim index As Integer = html.IndexOf("01_")
            If index <> -1 AndAlso index + 50 <= html.Length Then
                Dim substring As String = html.Substring(index, 50)
                MsgBox(substring)
            End If

            ' Return the full HTML content of the website
            Return html
        End Function



    End Class

End Class