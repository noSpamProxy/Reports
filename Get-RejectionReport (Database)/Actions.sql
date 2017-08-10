declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)

select 
	name, 
	case  
		when Decision = 2 then 'Temporary Blocked'
		when Decision = 3 then 'Permanently Blocked'
	end Decision, 
	count(*) Count 
	from MessageTracking.Action a
	join MessageTracking.MessageTrackEntry m
	on a.MessageTrackId = m.id
	where m.Sent >@startdate and m.Sent < @enddate and (m.status = 3 or m.status = 4) and (Decision = 2 or Decision = 3)
	group by name, Decision
