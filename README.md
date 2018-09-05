# vba-sms
envoie de sms en vba api FREE

````
Sub SendSMS()
    Dim strReturn As String
    Dim Pass As String
    Dim User As String
    Dim Message As String
    
    User = "utilisateur"
    Pass = "mot de passe"
    
    Message = "ici le message"
    strReturn = send(User, Pass, WorksheetFunction.EncodeURL(Message))
    Debug.Print strReturn
End Sub
```


```
Function send(User, Pass, Message) As String
    Dim objWinHTTP As Object
    Dim strReturn As String
    Dim Request As String
    Dim url As String

    Set objWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    url = "https://smsapi.free-mobile.fr/sendmsg?"
    Request = "user=" & User & "&pass=" & Pass
    Request = Request & "&msg=" & Message



    objWinHTTP.Open "GET", url & Request, False
    objWinHTTP.SetTimeouts 30000, 30000, 30000, 30000
    objWinHTTP.send
    If objWinHTTP.StatusText = "OK" Then
        strReturn = objWinHTTP.ResponseText
        Debug.Print strReturn
    End If
    Set objWinHTTP = Nothing
    send = strReturn
End Function
```
