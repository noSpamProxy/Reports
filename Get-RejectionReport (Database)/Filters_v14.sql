declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)
declare @tenantId int = '{2}';

select 
	name, 
	count(*) Count 
	from MessageTracking.Filter f
	join MessageTracking.MessageTrackEntry m
	on f.MessageTrackId = m.id AND m.TenantId = f.TenantId
	where m.Sent > @startdate and m.Sent < @enddate and (m.status = 3 or m.status = 4) and (f.Scl > 0) AND m.TenantId = @tenantId
	group by name