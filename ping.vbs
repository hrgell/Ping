Option Explicit

' Copied and modified from
' https://learn.microsoft.com/en-us/previous-versions/windows/desktop/wmipicmp/win32-pingstatus
'
' Command line
' cscript /NOLOGO ping2.vbs


Main

Function Main
    Dim IPName
    'If GetOsVersion < "5.01" Then
    '  MsgBox "Unsupported Operating System", vbOKOnly  + vbExclamation, WScript.ScriptFullName
    '  WScript.Quit
    'End If

    'IPName = InputBox ("Specify an address to ping", WScript.ScriptFullName, "google.com")
    IPName = "telenor.de"
    Do
        PingTest IPName
        WScript.Sleep 2000
    Loop
End Function

Function PingTest (IPName)
    Dim Ping, Success, Status
    Set Ping = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & IPName & "'")
    If IsNull(Ping.StatusCode) Or IsEmpty(Ping.StatusCode) Then
        Success = False
    Else
        Success = (Ping.StatusCode = 0)
    End If
    If not Success Then
        Status = GetStatusCodeDescription(Ping.StatusCode)
        WScript.Echo Now & ": " & IPName & ": Failed (" & Status & ")"
        EchoPingInfo Ping
    'Else
    '    Status = GetStatusCodeDescription(Ping.StatusCode)
    '    WScript.Echo Now & ": " & IPName & ": (" & Status & ")"
    '   EchoPingInfo Ping
    End If
End Function

Function EchoPingInfo (Ping)
    With Ping
        Wscript.Echo "Address : " & .Address
        Wscript.Echo "Buffer size : " & .BufferSize
        Wscript.Echo "No Fragmentation : " & .NoFragmentation
        Wscript.Echo "PrimaryAddressResolutionStatus : " & .PrimaryAddressResolutionStatus
        Wscript.Echo "ProtocolAddress : " & .ProtocolAddress
        Wscript.Echo "ProtocolAddressResolved : " & .ProtocolAddressResolved
        Wscript.Echo "RecordRoute : " & .RecordRoute
        Wscript.Echo "ReplyInconsistency : " & .ReplyInconsistency
        Wscript.Echo "ReplySize : " & .ReplySize
        Wscript.Echo "ResolveAddressNames : " & .ResolveAddressNames
        Wscript.Echo "ResponseTime : " & .ResponseTime
        Wscript.Echo "ResponseTimeToLive : " & .ResponseTimeToLive
        If IsNull (.RouteRecord) Then
          Wscript.Echo "RouteRecord : Null"
        Else
          Wscript.Echo "RouteRecord : " & _
                 Join (.RouteRecord, "; ")
        End If
        If IsNull (.RouteRecordResolved) Then
          Wscript.Echo "RouteRecordResolved : Null"
        Else
          Wscript.Echo "RouteRecordResolved : " & _
                 Join (.RouteRecordResolved, "; ")
        End If
        Wscript.Echo "SourceRoute : " & .SourceRoute
        Wscript.Echo "SourceRouteType : " & GetSourceRouteType(.SourceRouteType)
        Wscript.Echo "Status code : " & GetStatusCodeDescription(.StatusCode)
        Wscript.Echo "Timeout : " & .TimeOut
        If IsNull (.TimeStampRecord) Then
          Wscript.Echo "TimeStampRecord : Null"
        Else
          Wscript.Echo "TimeStampRecord : " & _
                 Join (.TimeStampRecord, "; ")
        End If
        If IsNull (.TimeStampRecordAddress) Then
          Wscript.Echo "TimeStampRecordAddress : Null"
        Else
          Wscript.Echo "TimeStampRecordAddress : " & _
                 Join (.TimeStampRecordAddress, "; ")
        End If
        If IsNull (.TimeStampRecordAddressResolved) Then
          Wscript.Echo "TimeStampRecordAddressResolved : Null"
        Else
          Wscript.Echo "TimeStampRecordAddressResolved : " & _
                 Join (.TimeStampRecordAddressResolved, "; ")
        End If
        Wscript.Echo "TimeStampRoute : " & .TimeStampRoute
        Wscript.Echo "TimeToLive : " & .TimeToLive
        Wscript.Echo "TypeOfService : " & GetTypeOfService(.TypeOfService)
        Wscript.Echo ""
    End With
End Function

Function GetOsVersion
    Dim WMIService, Items, Element
    Set WMIService = GetObject("winmgmts:") 
    Set Items = WMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
    For Each Element In Items
        WScript.Echo Element.Version
        GetOsVersion = Element.Version
    Next
    Set Items = Nothing
    Set WMIService = Nothing
End Function

Function GetSourceRouteType (SourceRouteType)
    Dim Txt
    Select Case SourceRouteType
    Case 1
        Txt = "Loose Source Routing"
    Case 2
        Txt = "Strict Source Routing"
    Case Else
        ' Default - 0 - or any other value.
        Txt = SourceRouteType & " - None"
    End Select
    GetSourceRouteType = Txt
End Function

Function GetTypeOfService (ServiceType)
    Dim Txt
    Select Case ServiceType
    Case 2
        Txt = "Minimize Monetary Cost"
    Case 4
        Txt = "Maximize Reliability"
    Case 8
        Txt = "Maximize Throughput"
    Case 16
        Txt = "Minimize Delay"
    Case Else
        ' Default - 0 - or any other value.
        Txt = ServiceType & " - Normal"
    End Select
    GetTypeOfService = Txt
End Function

Function GetStatusCodeDescription (StatusCode)
    Dim Txt
    If IsNull(StatusCode) Or IsEmpty(StatusCode) Then
        GetStatusCodeDescription = "Unknown"
        Exit Function
    End If
    Select Case StatusCode
    Case 0
        Txt = "Success"
    Case 11001
        Txt = "Buffer Too Small"
    Case 11002
        Txt = "Destination Net Unreachable"
    Case 11003
        Txt = "Destination Host Unreachable"
    Case 11004
        Txt = "Destination Protocol Unreachable"
    Case 11005
        Txt = "Destination Port Unreachable"
    Case 11006
        Txt = "No Resources"
    Case 11007
        Txt = "Bad Option"
    Case 11008
        Txt = "Hardware Error"
    Case 11009
        Txt = "Packet Too Big"
    Case 11010
        Txt = "Request Timed Out"
    Case 11011
        Txt = "Bad Request"
    Case 11012
        Txt = "Bad Route"
    Case 11013
        Txt = "TimeToLive Expired Transit"
    Case 11014
        Txt = "TimeToLive Expired Reassembly"
    Case 11015
        Txt = "Parameter Problem"
    Case 11016
        Txt = "Source Quench"
    Case 11017
        Txt = "Option Too Big"
    Case 11018
        Txt = "Bad Destination"
    Case 11032
        Txt = "Negotiating IPSEC"
    Case 11050
        Txt = "General Failure"
    Case Else
        Txt = StatusCode & " - Unknown"
    End Select
    GetStatusCodeDescription = Txt
End Function
