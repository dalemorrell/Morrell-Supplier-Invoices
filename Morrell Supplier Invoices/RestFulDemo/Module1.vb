Imports System.IO
Imports System.Net
Imports System.Text

Module Module1

    Sub Main()
        Do
            Try
                Dim content As String
                Console.WriteLine("Enter Method:")
                Dim Method = Console.ReadLine().ToUpper
                Console.WriteLine("Enter URI:")
                Dim Uri = Console.ReadLine()
                Dim req As HttpWebRequest = WebRequest.Create(Uri)
                req.KeepAlive = False
                req.Method = Method

                If ({"POST,PUT"}.Contains(Method)) Then
                    Console.WriteLine("Enter XML FilePath:")
                    Dim FilePath = Console.ReadLine()
                    content = (File.OpenText(FilePath)).ReadToEnd()
                    Dim buffer = Encoding.ASCII.GetBytes(content)
                    req.ContentLength = buffer.Length
                    req.ContentType = "text/xml"
                    Dim PostData = req.GetRequestStream()
                    PostData.Write(buffer, 0, buffer.Length)
                    PostData.Close()
                End If

                Dim resp As HttpWebResponse = req.GetResponse()
                Dim enc = System.Text.Encoding.GetEncoding(1252)
                Dim loResponseStream = New StreamReader(resp.GetResponseStream(), enc)
                Dim Response = loResponseStream.ReadToEnd()
                loResponseStream.Close()
                resp.Close()
                Console.WriteLine(Response)

            Catch ex As Exception
                Console.WriteLine(ex.Message.ToString())
            End Try
            Console.WriteLine()
            Console.WriteLine("Do you want to continue?")
        Loop Until Console.ReadKey().ToString.ToUpper() = "Y"
    End Sub

End Module
