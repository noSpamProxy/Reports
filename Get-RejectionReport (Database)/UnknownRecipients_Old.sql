declare @startdate datetime = convert(datetime, '{0}', 104)
declare @enddate datetime = convert(datetime, '{1}', 104)

select count(*) from messagetracking.messagetrackentry mt
	where WasReceivedFromRelayServer = 0 and  Sent > @startdate and Sent < @enddate and not exists (
		select 1
		from messagetracking.deliveryattempt da
		where da.MessageTrackId = mt.id and (da.Status <> 4 or da.StatusMessage <> N'Recipient unknown' OR da.StatusMessage is null)
		)

