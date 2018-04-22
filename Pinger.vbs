Option Explicit

Dim strComputer, objWMIService, colItems, objItem

On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

strComputer = "192.168.1.101"

Set objWMIService = GetObject("winmgmts:\\")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus Where Address = '" & strComputer & "'")


For Each objItem In colItems
   WScript.Echo "Address: " & objItem.Address
   WScript.Echo "BufferSize: " & objItem.BufferSize
   WScript.Echo "NoFragmentation: " & objItem.NoFragmentation
   WScript.Echo "PrimaryAddressResolutionStatus: " & objItem.PrimaryAddressResolutionStatus
   WScript.Echo "ProtocolAddress: " & objItem.ProtocolAddress
   WScript.Echo "ProtocolAddressResolved: " & objItem.ProtocolAddressResolved
   WScript.Echo "RecordRoute: " & objItem.RecordRoute
   WScript.Echo "ReplyInconsistency: " & objItem.ReplyInconsistency
   WScript.Echo "ReplySize: " & objItem.ReplySize
   WScript.Echo "ResolveAddressNames: " & objItem.ResolveAddressNames
   WScript.Echo "ResponseTime: " & objItem.ResponseTime
   WScript.Echo "ResponseTimeToLive: " & objItem.ResponseTimeToLive
   strRouteRecord = Join(objItem.RouteRecord, ",")
      WScript.Echo "RouteRecord: " & strRouteRecord
   strRouteRecordResolved = Join(objItem.RouteRecordResolved, ",")
      WScript.Echo "RouteRecordResolved: " & strRouteRecordResolved
   WScript.Echo "SourceRoute: " & objItem.SourceRoute
   WScript.Echo "SourceRouteType: " & objItem.SourceRouteType
   WScript.Echo "StatusCode: " & objItem.StatusCode
   WScript.Echo "Timeout: " & objItem.Timeout
   strTimeStampRecord = Join(objItem.TimeStampRecord, ",")
      WScript.Echo "TimeStampRecord: " & strTimeStampRecord
   strTimeStampRecordAddress = Join(objItem.TimeStampRecordAddress, ",")
      WScript.Echo "TimeStampRecordAddress: " & strTimeStampRecordAddress
   strTimeStampRecordAddressResolved = Join(objItem.TimeStampRecordAddressResolved, ",")
      WScript.Echo "TimeStampRecordAddressResolved: " & strTimeStampRecordAddressResolved
   WScript.Echo "TimestampRoute: " & objItem.TimestampRoute
   WScript.Echo "TimeToLive: " & objItem.TimeToLive
   WScript.Echo "TypeofService: " & objItem.TypeofService
   WScript.Echo
Next

