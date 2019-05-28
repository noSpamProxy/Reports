WITH Query AS (
	SELECT
		s.Domain,
		COALESCE (u.Displayname, s.Address) AS DisplayName, 
			CASE WHEN SUM(s.MailsSent) = 0 THEN 0 ELSE 1 END AS Protection,
			CASE WHEN SUM((s.SMimeMailsSigned + s.SMimeMailsEncrypted + s.PgpMailsSigned + s.PgpMailsEncrypted + s.PdfMailsSent)) = 0 THEN 0 ELSE 1 END AS [Encryption], 
            CASE WHEN SUM(s.MailsWithDisclaimer) = 0 THEN 0 ELSE 1 END AS Disclaimer, 
			CASE WHEN SUM(s.MailsWithLargeFiles) = 0 THEN 0 ELSE 1 END AS LargeFiles, 
			SUM(CASE WHEN DATEDIFF(d, s.Date,GetDate()) > 30 THEN 0 ELSE s.FilesUploadedToSandbox END) AS FilesUploadedToSandbox
	FROM MessageTracking.UserAndDomainStatistic AS s 
		LEFT OUTER JOIN	Usermanagement.MailAddress AS ma ON ma.MailAddress = s.Address 
		LEFT OUTER JOIN	Usermanagement.[User] AS u ON u.Id = ma.UserId
    WHERE(s.Date > DATEADD(d, - 90, GETDATE()))
    GROUP BY COALESCE (u.Displayname, s.Address), s.Domain
)
SELECT
	Domain,
	DisplayName, 
	Protection, 
	[Encryption], 
	Disclaimer, 
	LargeFiles, 
	FilesUploadedToSandbox
FROM Query
WHERE (Protection <> 0) 
	OR ([Encryption] <> 0) 
	OR (Disclaimer <> 0) 
	OR (LargeFiles <> 0) 
	OR (FilesUploadedToSandbox <> 0)
