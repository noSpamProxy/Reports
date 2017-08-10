declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)

select 
	case 
		when WasReceivedFromRelayServer = 0 then 'Inbound'
		when WasReceivedFromRelayServer = 1 then 'Outbound'
		else 'Summary'
	end Direction, 
	case  
		when Status = 1 then 'Success' 
		when Status = 3 then 'Temporary Blocked'
		when Status = 4 then 'Permanently Blocked'
		when Status is null then 'Summary'
	end Status, 
	COUNT(*)  Count
from MessageTracking.MessageTrackEntry
where Sent > @startdate and Sent < @enddate and (status = 1 or status = 3 or status = 4)
group by rollup (WasReceivedFromRelayServer, status)