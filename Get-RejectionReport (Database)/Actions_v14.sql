declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)
declare @tenantId int = '{2}';

select 
	name, 
	case  
		when Decision = 2 then 'Temporary Blocked'
		when Decision = 3 then 'Permanently Blocked'
	end Decision, 
	count(*) Count 
	from MessageTracking.Action a
	join MessageTracking.MessageTrackEntry m
	on a.MessageTrackId = m.id AND a.TenantId = m.TenantId
	where m.Sent >@startdate and m.Sent < @enddate and m.status IN (3,4) and Decision IN (2,3) AND m.TenantId = @tenantId
	group by name, Decision
