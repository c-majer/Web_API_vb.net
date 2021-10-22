Imports System
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports Newtonsoft.Json
Imports Json

Module Program
    Sub Main(args As String())
        Dim con As ADODB.Connection
        Dim sql As String
        Dim rs As ADODB.Recordset

        con = New ADODB.Connection
        con.ConnectionString = "Provider='###';data source=###; Initial Catalog=###; User Id=###; Password='###';"
        con.Open()

        sql = "Select ###"
        rs = New ADODB.Recordset
        rs.Open(sql, con, ADODB.CursorTypeEnum.adOpenDynamic)

        Dim daylist As New List(Of Product)

        While Not rs.EOF 

            Dim product As New Product()
            product.timestamp = rs(0).Value.ToString()
            product.sesso = rs(1).Value.ToString()
            product.titStu = rs(2).Value.ToString()
            product.cittadinanza = rs(3).Value.ToString()
            product.telPz = rs(4).Value.ToString()
            product.emailPz = rs(5).Value.ToString()
            product.datanasc = rs(6).Value.ToString()
            product.codcomres = rs(7).Value.ToString()
            product.aslres = rs(8).Value.ToString()
            product.codpo = rs(9).Value.ToString()
            product.codrepamm = rs(10).Value.ToString()
            product.codrepdim = rs(11).Value.ToString()
            product.datamm = rs(12).Value.ToString()
            product.datadim = rs(13).Value.ToString()
            product.tras1 = rs(14).Value.ToString()
            product.rep1 = rs(15).Value.ToString()
            product.tras2 = rs(16).Value.ToString()
            product.rep2 = rs(17).Value.ToString()
            product.tras3 = rs(18).Value.ToString()
            product.rep3 = rs(19).Value.ToString()
            product.token = rs(20).Value.ToString()

            daylist.Add(product)

            rs.MoveNext()

        End While

        Dim jsonFile As String
        jsonFile = JsonConvert.SerializeObject(daylist)

        rs.Close()
        rs = Nothing
        con.Close()
        con = Nothing

    End Sub

        Public siteUri As New Uri("Web Endpoint URL")
        Public JavaScriptConvert As Object

    Public Function UploadValues(siteUri As Uri, POST As String, jsonFile As JsonDocument)
        Return jsonFile
    End Function

End Module

Friend Class Product
    
    Public Property timestamp As String
    Public Property sesso As String
    Public Property titStu As String
    Public Property cittadinanza As String
    Public Property telPz As String
    Public Property emailPz As String
    Public Property datanasc As String
    Public Property codcomres As String
    Public Property aslres As String
    Public Property codpo As String
    Public Property codrepamm As String
    Public Property codrepdim As String
    Public Property datamm As String
    Public Property datadim As String
    Public Property tras1 As String
    Public Property rep1 As String
    Public Property tras2 As String
    Public Property rep2 As String
    Public Property tras3 As String
    Public Property rep3 As String
    Public Property token As String

End Class
