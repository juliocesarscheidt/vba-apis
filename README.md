# VBA requests para chamar APIs externas

## APIs

```bash
#### coinbase ####
# https://developers.coinbase.com/docs/wallet/guides/price-data
# https://developers.coinbase.com/api/v2#get-spot-price

# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'https://api.coinbase.com/v2/prices/BTC-BRL/spot'


#### biscoint ####
# https://biscoint.io/docs/api#public-ticker

# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'https://api.biscoint.io/v1/ticker?base=BTC&quote=BRL'


#### bitpreco ####
# https://bitpreco.com/api.html

# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'https://api.bitpreco.com/btc-brl/ticker'


#### awesomeapi ####
# https://docs.awesomeapi.com.br/api-de-moedas

# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'https://economia.awesomeapi.com.br/all/USD-BRL'


#### viacep ####
# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'http://viacep.com.br/ws/80420010/json'


#### geo qualaroo ####
# exemplo com curl
curl --silent -H 'Content-Type: application/json' \
  'https://geo.qualaroo.com/json/'
```

## VBA

```vb
Public Function GET_LATITUDE() As String
    Dim apiUrl$
    apiUrl = "https://geo.qualaroo.com/json/"

    rawResult = REQUEST(apiUrl, "GET")
    GET_LATITUDE = handleJsonParts(rawResult, "latitude")
End Function

Public Function GET_LONGITUDE() As String
    Dim apiUrl$
    apiUrl = "https://geo.qualaroo.com/json/"

    rawResult = REQUEST(apiUrl, "GET")
    GET_LONGITUDE = handleJsonParts(rawResult, "longitude")
End Function

Public Function GET_ZIP_CODE_INFO(ByVal zipCode$, ByVal fieldName$) As String
    Dim apiUrl$
    apiUrl = "http://viacep.com.br/ws/"

    rawResult = REQUEST(apiUrl & zipCode & "/json", "GET")
    GET_ZIP_CODE_INFO = handleJsonParts(rawResult, fieldName)
End Function

Public Function GET_CURRENCY_TICKET(Optional ByVal currencyKind$, Optional ByVal ticketKind$) As Currency
    Dim apiUrl$, finalResult$
    apiUrl = "https://economia.awesomeapi.com.br/all/"

    'currencyKind default is "USD-BRL"
    'accepted values are: "USD-BRL", "EUR-BRL", "GBP-BRL", "BTC-BRL", etc
    If currencyKind = "" Then
        currencyKind = "USD-BRL"
    End If

    'ticketKind default is "bid"
    'accepted values are: "bid", "ask", "high", "low"
    If ticketKind = "" Then
        ticketKind = "bid"
    End If

    rawResult = REQUEST(apiUrl & currencyKind, "GET")
    finalResult = handleJsonParts(rawResult, ticketKind)

    GET_CURRENCY_TICKET = CCur(Replace(finalResult, ".", ","))
End Function

Public Function GET_BITCOIN_TICKET(Optional ByVal ticketKind$) As Currency
    Dim apiUrl$, finalResult$
    apiUrl = "https://api.bitpreco.com/btc-brl/ticker"

    'ticketKind default is "last"
    'accepted values are: "last", "buy", "sell", "high", "low"
    If ticketKind = "" Then
        ticketKind = "last"
    End If

    rawResult = REQUEST(apiUrl, "GET")
    finalResult = handleJsonParts(rawResult, ticketKind)

    GET_BITCOIN_TICKET = CCur(Replace(finalResult, ".", ","))
End Function

Private Function handleJsonParts(ByVal jsonDataString$, ByVal fieldName$) As String
    Dim finalResult$
    Dim jsonParts() As String
    Dim jsonFields() As String

    'turn json string into parts split by comma
    jsonParts = Split(jsonDataString, ",")

    For Each jsonPart In jsonParts
        jsonPart = Replace(jsonPart, "{", "")
        jsonPart = Replace(jsonPart, "}", "")
        jsonPart = Replace(jsonPart, """", "") 'remove double quotation marks

        jsonFields = Split(jsonPart, ":")
        jsonFields(0) = Trim(jsonFields(0))
        jsonFields(1) = Trim(jsonFields(1))

        If jsonFields(0) = fieldName Then
            finalResult = jsonFields(1)
            Exit For 'break
        End If
    Next

    handleJsonParts = finalResult
End Function

Public Function REQUEST(ByVal apiUrl$, ByVal method$, Optional ByVal jsonDataString$, Optional ByVal bearerToken$) As String
    Dim objHTTP As Object
    Dim responseCode$, responseText$

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objHTTP.Open method, apiUrl, False
    objHTTP.setRequestHeader "Content-type", "application/json"

    'setting oauth token when provided
    If bearerToken <> "" Then
        objHTTP.setRequestHeader "Authorization", "Bearer " & bearerToken
    End If

    'setting payload when provided
    If jsonDataString <> "" Then
        objHTTP.Send (jsonDataString)
    Else
        objHTTP.Send
    End If

    responseCode = objHTTP.Status

    If responseCode >= 200 And responseCode <= 299 Then
        responseText = objHTTP.responseText

        responseText = Replace(responseText, Chr(10), "")
        responseText = Replace(responseText, Chr(13), "")
    End If

    Set objHTTP = Nothing

    REQUEST = responseText
End Function
```
