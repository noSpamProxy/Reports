declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)

select
	COUNT(*) AS Count,
	MessageTracking.MessageAddress.Address,
	MessageTracking.MessageAddress.Domain
from MessageTracking.MessageTrackEntry
	join MessageTracking.MessageAddress
		on MessageTracking.MessageAddress.MessageTrackId = MessageTracking.MessageTrackEntry.Id
where Sent > @startdate and Sent < @enddate and (status = 3 or status = 4) and MessageTracking.MessageAddress.AddressType = 2
group by Address, Domain, AddressType
order by Count desc