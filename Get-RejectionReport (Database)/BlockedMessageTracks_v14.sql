declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)
declare @tenantId int = '{2}';

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
		when Status = 6 then 'PartialSuccess'
		when Status = 9 then 'Put on Hold'
		when Status is null then 'Summary'
	end Status,
	COUNT(*)  Count
from MessageTracking.MessageTrackEntry 
where Sent > @startdate and Sent < @enddate and status IN (1,3,4,6,9) AND TenantId = @tenantId
group by rollup (WasReceivedFromRelayServer, status)