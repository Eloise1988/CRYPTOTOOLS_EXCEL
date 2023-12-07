Attribute VB_Name = "CryptoTools"
''
' CRYPTOTOOLS v1.0.3
' (c) Eloise1988 - https://github.com/Eloise1988/CRYPTOTOOLS_EXCEL
'
' Cryptocurrency Market Data
'
' @class CryptoTools
' @author ac@charmantadvisory.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on cryptotools Google Sheet version (with extensive changes)
'
' CryptoTools.gs, https://github.com/Eloise1988/CRYPTOBALANCE/blob/master/CRYPTOTOOLS_V2.gs
'
' Copyright (c) 2023, Eloise1988
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

Public Const CRYPTOTOOLS_API_KEY As String = "my_api_key"

Sub AddFunctionDescription()
' Adds a description to cryptotools functions
    Application.MacroOptions Macro:="CRYPTOPRICE", Description:="Returns cryptocurrency prices in USD"
    Application.MacroOptions Macro:="CRYPTOBALANCE", Description:="Returns cryptocurrency wallet balances"
    Application.MacroOptions Macro:="CRYPTONETWORTH", Description:="Returns a wallet's networth in USD"
    Application.MacroOptions Macro:="CRYPTODEXPRICE", Description:="Returns a list of prices from specific decentralized exchanges"
    Application.MacroOptions Macro:="CRYPTOVOLATILITY", Description:="Returns cryptocurrency 30 Day volatility against USD"
End Sub
Public Function CRYPTOPRICE(ticker As Variant) As Variant
Attribute CRYPTOPRICE.VB_Description = "Returns cryptocurrency prices in USD"
Attribute CRYPTOPRICE.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns cryptocurrency prices in USD
' ticker: The array of tickers you want prices from
    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim CallerRows As Long
    Dim Field As String
    Dim k As Long
    Application.Calculation = xlCalculationManual
    
    ' Set default values
    Field = "PRICE"
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If
    
    If TypeOf ticker Is Range Then
        ' Set default values
        CallerRows = ticker.Rows.Count
        
        ' Construct API URL
        URL = "/CRYPTOPRICE/" & ticker(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & ticker(k, 1).value
        Next k
      
    Else
        ' Construct API URL
        URL = "/CRYPTOPRICE/" & ticker
        
    End If
    
    
    URL = URL & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send
    
    ' Parse JSON response
    Dim json As Object
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    Dim output() As Variant
    ReDim Preserve output(1 To json.Count, 0)
    
    ' Extract field value from each JSON object
    Dim i As Long
    For i = 1 To json.Count
        If Not json(i).Exists(Field) Then
            output(i, 0) = ""
        End If
        output(i, 0) = Val(json(i)(Field))
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Return output array
    CRYPTOPRICE = output
    
End Function

Public Function CRYPTOBALANCE(ticker As Variant, address As Variant) As Variant
Attribute CRYPTOBALANCE.VB_Description = "Returns cryptocurrency wallet balances"
Attribute CRYPTOBALANCE.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns cryptocurrency wallet balances
' ticker: The array of tickers you want balances from
' address: The array of wallet addresses you want balances from
    
    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim CallerRows As Long
    Dim Field As String
    Dim k As Long
    
    Application.Calculation = xlCalculationManual
    
    Field = "QUANTITY"
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If
    
    
    If TypeOf ticker Is Range Then
        ' Set default values
        CallerRows = ticker.Rows.Count
        
        
        ' Construct API URL
        URL = "/BALANCES/" & ticker(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & ticker(k, 1).value
        Next k
        URL = URL & "/" & address(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & address(k, 1).value
        Next k
      
    Else
        ' Construct API URL
        URL = "/BALANCES/" & ticker & "/" & address
        
    End If
    
    URL = URL & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send
    
    ' Parse JSON response
    Dim json As Object
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    Dim output() As Variant
    ReDim Preserve output(1 To json.Count, 0)
    
    ' Extract field value from each JSON object
    Dim i As Long
    For i = 1 To json.Count
        If Not json(i).Exists(Field) Then
            output(i, 0) = ""
        End If
        output(i, 0) = Val(json(i)(Field))
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Return output array
    CRYPTOBALANCE = output
    
End Function


Public Function CRYPTONETWORTH(address As Variant) As Variant
Attribute CRYPTONETWORTH.VB_Description = "Returns a wallet's networth in USD"
Attribute CRYPTONETWORTH.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns a wallet's networth in USD
' address: The wallet addresse you want the networth sum from
    
    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim CallerRows As Long
    
    Application.Calculation = xlCalculationManual
    
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If
    
    
    If TypeOf address Is Range Then
        ' Set default values
        CallerRows = address.Rows.Count
        
        
        ' Construct API URL
        URL = "/TOTALUSDBALANCE/" & address(1, 1).value
       
      
    Else
        ' Construct API URL
        URL = "/TOTALUSDBALANCE/" & address
        
    End If
    
    URL = URL & "/ALL/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Return output array
    CRYPTONETWORTH = Val(request.responseText)
    
    
End Function

Public Function CRYPTODEXPRICE(token1 As Variant, token2 As Variant, exchange As Variant) As Variant
Attribute CRYPTODEXPRICE.VB_Description = "Returns a list of prices from specific decentralized exchanges"
Attribute CRYPTODEXPRICE.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns cryptocurrency dex prices
' token1: The array of token1 from pair (token1/token2)
' token2: The array of token2 from pair (token1/token2)
' exchange: The array of dex exchanges for prices
    
    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim CallerRows As Long
    Dim Field As String
    Dim Field1 As String
    Dim k As Long
    
    Application.Calculation = xlCalculationManual
    
    Field = "PRICE"
    Field1 = "DEXPRICE2"
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If
    
    
    If TypeOf token1 Is Range Then
        ' Set default values
        CallerRows = token1.Rows.Count
        
        
        ' Construct API URL
        URL = "/" & Field1 & "/" & token1(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & token1(k, 1).value
        Next k
        URL = URL & "/" & token2(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & token2(k, 1).value
        Next k
        URL = URL & "/" & exchange(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & exchange(k, 1).value
        Next k
      
    Else
        ' Construct API URL
        URL = "/" & Field1 & "/" & token1 & "/" & token2 & "/" & exchange
        
    End If
    
    URL = URL & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send
    
    ' Parse JSON response
    Dim json As Object
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    Dim output() As Variant
    ReDim Preserve output(1 To json.Count, 0)
    
    ' Extract field value from each JSON object
    Dim i As Long
    For i = 1 To json.Count
        If Not json(i).Exists(Field) Then
            output(i, 0) = ""
        End If
        output(i, 0) = Val(json(i)(Field))
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Return output array
    CRYPTODEXPRICE = output
    
End Function
Public Function CRYPTOHIST(ticker As Variant, datatype As String, startdate As String, enddate As String) As Variant
Attribute CRYPTOHIST.VB_Description = "Returns the historical cryptocurrency OHLC data"
Attribute CRYPTOHIST.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns the historical cryptocurrency OHLC data
' ticker: Array of tickers (max 3 on freemium)
' datatype: "open", "high", "low", "close", "volume", "marketcap"
' startdate: Start date in "yyyy-mm-dd" format
' enddate: End date in "yyyy-mm-dd" format

    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim json As Object
    Dim data() As Variant
    Dim i As Long, k As Long
    Dim CallerRows As Long
    Application.Calculation = xlCalculationManual
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If

    ' Construct API URL
    If TypeOf ticker Is Range Then
        CallerRows = ticker.Rows.Count
        URL = "/PRICEHISTO/" & ticker(1, 1).Value
        For k = 2 To CallerRows
            URL = URL & "%2C" & ticker(k, 1).Value
        Next k
    Else
        URL = "/PRICEHISTO/" & ticker
    End If

    URL = URL & "/" & datatype & "/" & startdate & "/" & enddate & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send

    ' Parse JSON response
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    ReDim Preserve data(1 To json.Count, 1 To 5) ' Assuming 5 fields: Open, High, Low, Close, Volume

    ' Extract field value from each JSON object
    For i = 1 To json.Count
        data(i, 1) = Val(json(i)("OPEN"))
        data(i, 2) = Val(json(i)("HIGH"))
        data(i, 3) = Val(json(i)("LOW"))
        data(i, 4) = Val(json(i)("CLOSE"))
        data(i, 5) = Val(json(i)("VOLUME"))
    Next i

    Application.Calculation = xlCalculationAutomatic

    ' Return output array
    CRYPTOHIST = data

End Function
Public Function CRYPTOTOKENLIST(address As String, Optional chain As String = "all") As Variant
Attribute CRYPTOTOKENLIST.VB_Description = "Returns the list of all tokens on specified chain or all chains"
Attribute CRYPTOTOKENLIST.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns the list of all tokens on specified chain or all chains
' address: The wallet address
' chain: The blockchain to query (default is "all")

    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim json As Object
    Dim i As Long
    Dim data() As Variant
    Application.Calculation = xlCalculationManual
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If

    ' Construct API URL
    URL = "/CRYPTOLIST/" & address & "/" & chain & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send

    ' Parse JSON response
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    ReDim Preserve data(1 To json.Count, 1 To 6) ' Assuming 6 fields: CHAIN, CONTRACT, SYMBOL, QTY, PRICE, 'AMOUNT ($)'
    
    ' Extract field values from each JSON object
    For i = 1 To json.Count
        data(i, 1) = json(i)("CHAIN")
        data(i, 2) = json(i)("CONTRACT")
        data(i, 3) = json(i)("SYMBOL")
        data(i, 4) = Val(json(i)("QTY"))
        data(i, 5) = Val(json(i)("PRICE"))
        data(i, 6) = Val(json(i)("AMOUNT ($)"))
    Next i

    Application.Calculation = xlCalculationAutomatic

    ' Return output array
    CRYPTOTOKENLIST = data

End Function
Public Function CRYPTOTX(addresses As Variant, network As String) As Variant
Attribute CRYPTOTX.VB_Description = "Returns the historical transaction list on a range of addresses"
Attribute CRYPTOTX.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns the historical transaction list on a range of addresses
' addresses: Array of addresses (max 3 on freemium)
' network: Available networks (e.g., btc, eth, bnb, etc.)

    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim json As Object
    Dim data() As Variant
    Dim i As Long, k As Long
    Dim CallerRows As Long
    Application.Calculation = xlCalculationManual

    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()

    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If

    ' Construct API URL
    If TypeOf addresses Is Range Then
        CallerRows = addresses.Rows.Count
        URL = "/TXALL/" & addresses(1, 1).Value
        For k = 2 To CallerRows
            URL = URL & "%2C" & addresses(k, 1).Value
        Next k
    Else
        URL = "/TXALL/" & addresses
    End If

    URL = URL & "/" & network & "/" & http_options("apikey")

    ' Combine private path and API URL
    URL = private_path & URL

    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send

    ' Parse JSON response
    Set json = JsonConverter.ParseJson(request.responseText)

    ' Create output array
    ReDim Preserve data(1 To json.Count, 1 To 5) ' Assuming 5 fields: transactional data fields

    ' Extract field value from each JSON object
    For i = 1 To json.Count
        data(i, 1) = json(i)("Field1") ' Replace "Field1" with actual field name
        data(i, 2) = json(i)("Field2") ' Replace "Field2" with actual field name
        ' ... continue for other fields as per the JSON structure
    Next i

    Application.Calculation = xlCalculationAutomatic

    ' Return output array
    CRYPTOTX = data

End Function

Public Function CRYPTOVOLATILITY(token As Variant) As Variant
Attribute CRYPTOVOLATILITY.VB_Description = "Returns cryptocurrency 30 Day volatility against USD"
Attribute CRYPTOVOLATILITY.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns cryptocurrency 30 Day volatility against USD
' token: The array of tokens you need the 30 day volatility from
    
    ' Declare variables and objects
    Dim URL As String
    Dim request As Object
    Dim private_path As String
    Dim http_options As Object
    Dim CallerRows As Long
    Dim Field As String
    Dim Field1 As String
    Dim k As Long
    
    Application.Calculation = xlCalculationManual
    
    Field = "VOLATILTY_30D"
    Field1 = "30DVOL"
    
    ' Set API endpoint and options
    private_path = "https://api.charmantadvisory.com"
    Set http_options = CreateObject("Scripting.Dictionary")
    http_options("apikey") = GetMyIPAddress()
    
    ' Check if custom API key is provided
    If CRYPTOTOOLS_API_KEY <> "my_api_key" Then
        private_path = "https://privateapi.charmantadvisory.com"
        Set http_options = CreateObject("Scripting.Dictionary")
        http_options.Add "headers", CreateObject("Scripting.Dictionary")
        http_options("headers")("apikey") = CRYPTOTOOLS_API_KEY
    End If
    
    
    If TypeOf token Is Range Then
        ' Set default values
        CallerRows = token.Rows.Count
        
        
        ' Construct API URL
        URL = "/" & Field1 & "/" & token(1, 1).value
        For k = 2 To CallerRows
            URL = URL & "%2C" & token(k, 1).value
        Next k
        URL = URL & "/USD"
        For k = 2 To CallerRows
            URL = URL & "%2CUSD"
        Next k
        
      
    Else
        ' Construct API URL
        URL = "/" & Field1 & "/" & token & "/USD"
        
    End If
    
    URL = URL & "/" & http_options("apikey")
    
    ' Combine private path and API URL
    URL = private_path & URL
    
    
    ' Send API request
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", URL, False
    request.setRequestHeader "apikey", http_options("apikey")
    request.send
    
    ' Parse JSON response
    Dim json As Object
    Set json = JsonConverter.ParseJson(request.responseText)
    
    ' Create output array
    Dim output() As Variant
    ReDim Preserve output(1 To json.Count, 0)
    
    ' Extract field value from each JSON object
    Dim i As Long
    For i = 1 To json.Count
        If Not json(i).Exists(Field) Then
            output(i, 0) = ""
        End If
        output(i, 0) = Val(json(i)(Field))
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    
    ' Return output array
    CRYPTOVOLATILITY = output
    
End Function
Function GetMyIPAddress() As String
    'Create a WinHttpRequest object using late binding
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim IP_API_URL As String
    
    IP_API_URL = "https://httpbin.org/ip"
    
    'Send an HTTP GET request to the httpbin.org/ip endpoint
    http.Open "GET", IP_API_URL, False
    http.send
    
    'Parse the JSON response and get the IP address
    Dim json As Object
    Dim IPAddr As String
    
    
    On Error Resume Next
    Set json = JsonConverter.ParseJson(http.responseText)
    If Err.Number <> 0 Then
        GetMyIPAddress = ""
        Exit Function
    End If
    On Error GoTo 0
    
    IPAddr = json("origin")
    
    'Convert the IP address to ASCII 256 format and return it
    GetMyIPAddress = ConvertToAscii256(IPAddr)
End Function

Function ConvertToAscii256(value As String) As String
    Dim i As Long ' Use Long data type for better performance and to avoid overflow issues for larger values
    Dim asciiVal As Integer
    Dim ascii256Val As String
    
    For i = 1 To Len(value)
        asciiVal = AscW(Mid(value, i, 1)) ' Use AscW instead of Asc to support Unicode characters
        ascii256Val = ascii256Val & Right$("000" & Hex$(asciiVal), 2) ' Use Right$ and Hex$ functions for better performance
    Next i
    
    ConvertToAscii256 = ascii256Val
End Function
