declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)

select 
	name, 
	count(*) Count 
	from MessageTracking.Filter f
	join MessageTracking.MessageTrackEntry m
	on f.MessageTrackId = m.id
	where m.Sent > @startdate and m.Sent < @enddate and (m.status = 3 or m.status = 4) and (f.Scl > 0)
	group by name