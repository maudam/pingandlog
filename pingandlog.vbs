' By Maudam
' https://github.com/maudam/pingandlog
' Initial release: jan, 20 2022
' V 1.0
' License: GplV2

' Reference for native win32 ping Wmi library:
' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus

' Hosts to ping:
strAddress1 = "google.com"
strAddress2 = "192.168.1.1"
strAddress3 = "192.168.1.200"
strAddress4 = "8.8.8.8"
strAddress5 = "1.1.1.1"

StartTime = Year(Date) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2) & "-" & Right("0" & hour(now),2) & Right("0" & minute(now),2)  & Right("0" & second(now),2)

wscript.echo StartTime & " Starting..."

Set fso = CreateObject("Scripting.FileSystemObject")
BuildFullPath "c:\windows\temp\PingAndLog"


' Logfile for every execution or unique file for all logs

'LogFile = "c:\windows\temp\Pinglogger\PingAndLog-" & StartTime & ".log"
LogFile = "c:\windows\temp\PingAndLog\PingAndLog.log"
wscript.echo LogFile

Set fso = CreateObject("Scripting.FileSystemObject" )

Mylog("Start ping....")

Do

  ' ping addresses. 500 ms pause on each ping e 5 second pause for next cycle
  ' Modify as needed
  pingaddress(strAddress1)
  WScript.Sleep 500
  pingaddress(strAddress2)
  WScript.Sleep 500
  pingaddress(strAddress3)
  WScript.Sleep 500
  pingaddress(strAddress4)
  WScript.Sleep 500
  pingaddress(strAddress5)
  WScript.Sleep 5000

Loop

' End of main script

' Functions & Subs

Function GetStatusCode (intCode)
  Dim strStatus
  Select Case intCode
  case  0
    strStatus = "Success"
  case  11001
    strStatus = "Buffer Too Small"
  case  11002
    strStatus = "Destination Net Unreachable"
  case  11003
    strStatus = "Destination Host Unreachable"
  case  11004
    strStatus = "Destination Protocol Unreachable"
  case  11005
    strStatus = "Destination Port Unreachable"
  case  11006
    strStatus = "No Resources"
  case  11007
    strStatus = "Bad Option"
  case  11008
    strStatus = "Hardware Error"
  case  11009
    strStatus = "Packet Too Big"
  case  11010
    strStatus = "Request Timed Out"
  case  11011
    strStatus = "Bad Request"
  case  11012
    strStatus = "Bad Route"
  case  11013
    strStatus = "TimeToLive Expired Transit"
  case  11014
    strStatus = "TimeToLive Expired Reassembly"
  case  11015
    strStatus = "Parameter Problem"
  case  11016
    strStatus = "Source Quench"
  case  11017
    strStatus = "Option Too Big"
  case  11018
    strStatus = "Bad Destination"
  case  11032
    strStatus = "Negotiating IPSEC"
  case  11050
    strStatus = "General Failure"
  case Else
    strStatus = intCode & " - Unknown"
  End Select
  GetStatusCode = strStatus
End Function


sub Mylog (msg)
  Set file = fso.OpenTextFile(LogFile,8,-1) 
  ActualTime = Year(Date) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2) & "-" & Right("0" & hour(now),2) & Right("0" & minute(now),2)  & Right("0" & second(now),2)
  File.Write(ActualTime & " " & msg & vbCrLf)
  wscript.echo msg
  File.Close
end sub


sub pingaddress(strAddress)
  Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strAddress & "'")
  With objPing
    ActualTime = Year(Date) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2) & "-" & Right("0" & hour(now),2) & Right("0" & minute(now),2)  & Right("0" & second(now),2)
    Wscript.Echo ActualTime & " : " & strAddress & " : " & .ProtocolAddress & " : " & GetStatusCode(.StatusCode)
    If GetStatusCode(.StatusCode) <> "Success" then
      Set file = fso.OpenTextFile(LogFile,8,-1)
      File.Write(ActualTime &  " : " & strAddress & " : " & .ProtocolAddress & " : " & GetStatusCode(.StatusCode) & vbCrLf)
      File.Close
    End if
  End With
End Sub

Sub BuildFullPath(ByVal FullPath)
    If Not fso.FolderExists(FullPath) Then
        BuildFullPath fso.GetParentFolderName(FullPath)
        fso.CreateFolder FullPath
    End If
End Sub
