Option Explicit

Dim A2, A1, s, f1, ts, ts1, fso, fso1, Soobshenie, SiteAddr1, SiteAddr2, bs, cur_time, api_key, api_secret, Message, HTTPgetCrypted
  Const ForReading = 1
  Const ForWriting = 2
 
Main

Sub Die()

End Sub



Sub Main

'NEXMO Credentials
api_key = "4432777a" 
api_secret = "cfg942576da14717"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile("nexmonumbers.txt", ForReading)
SiteAddr1 = "https://rest.nexmo.com/sms/json?api_key="&api_key&"&api_secret="&api_secret&"&from=12013743718&to="
SiteAddr2 = "&type=unicode&text="
Message = "This is a test message"

Do While Not ts.AtEndOfStream
    s = ts.ReadLine
    HTTPgetCrypted = SiteAddr1 & s & SiteAddr2 & Message 
    A1 = ""
    A1 = HTTPGetSend(HTTPGetCrypted)
    cur_time = time & " --- " & date
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    Set ts1 = fso1.OpenTextFile("nexmoresult.txt", 8, True)
    ts1.WriteLine(cur_time)		
    ts1.WriteLine("--------------------------------------------------------------------------------------")
    ts1.WriteLine(A1)
    ts1.Close
    Set fso1 = Nothing
    Set ts1 = Nothing
Loop
ts.Close
MsgBox "SMS sending complete, please see nexmoresult.txt for log"
WScript.Quit()
End Sub

Function HTTPGetSend(ByVal URLGet)
    Dim oHTTP
    Set oHTTP =  CreateObject("MSXML2.XMLHTTP")
    oHTTP.Open "GET",URLGet,False
    oHTTP.Send
    HTTPGetSend = oHTTP.responseText
    Set oHTTP = Nothing 
End Function
