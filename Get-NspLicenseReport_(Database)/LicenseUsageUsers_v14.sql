declare @tenantId int = '{0}';
WITH Query AS (
	SELECT
		COALESCE (u.Displayname, s.Address) AS DisplayName, 
			CASE WHEN SUM(s.MailsSent) = 0 THEN 0 ELSE 1 END AS Protection,
			CASE WHEN SUM((s.SMimeMailsSigned + s.SMimeMailsEncrypted + s.SMimeMailsDecrypted + s.SMimeMailsValidated + s.PgpMailsSigned + s.PgpMailsEncrypted + s.PgpMailsDecrypted + s.PgpMailsValidated + s.PdfMailsSent)) = 0 THEN 0 ELSE 1 END AS [Encryption], 
            CASE WHEN SUM(s.MailsWithDisclaimer) = 0 THEN 0 ELSE 1 END AS Disclaimer, 
			CASE WHEN SUM(s.MailsWithLargeFiles) = 0 THEN 0 ELSE 1 END AS LargeFiles, 
			CASE WHEN SUM(s.PdfMailsSent) = 0 THEN 0 ELSE 1 END AS PdfMailsSent, 
			CASE WHEN SUM(s.SMimeMailsSigned) = 0 THEN 0 ELSE 1 END AS SMimeMailsSigned, 
			CASE WHEN SUM(s.SMimeMailsEncrypted) = 0 THEN 0 ELSE 1 END AS SMimeMailsEncrypted, 
			CASE WHEN SUM(s.SMimeMailsValidated) = 0 THEN 0 ELSE 1 END AS SMimeMailsValidated, 
			CASE WHEN SUM(s.SMimeMailsDecrypted) = 0 THEN 0 ELSE 1 END AS SMimeMailsDecrypted, 
			CASE WHEN SUM(s.PgpMailsSigned) = 0 THEN 0 ELSE 1 END AS PgpMailsSigned, 
			CASE WHEN SUM(s.PgpMailsEncrypted) = 0 THEN 0 ELSE 1 END AS PgpMailsEncrypted, 
			CASE WHEN SUM(s.PgpMailsValidated) = 0 THEN 0 ELSE 1 END AS PgpMailsValidated, 
			CASE WHEN SUM(s.PgpMailsDecrypted) = 0 THEN 0 ELSE 1 END AS PgpMailsDecrypted, 
			SUM(CASE WHEN DATEDIFF(d, s.Date,GetDate()) > 30 THEN 0 ELSE s.FilesUploadedToSandbox END) AS FilesUploadedToSandbox
	FROM MessageTracking.UserAndDomainStatistic AS s
        LEFT OUTER JOIN Configuration.OwnedDomain AS od ON s.Domain = od.Name
        AND s.TenantId = od.TenantId
        LEFT OUTER JOIN Usermanagement.LocalAddress AS la ON la.Address = s.Address
        AND la.TenantId = s.TenantId
        LEFT OUTER JOIN Usermanagement.[User] AS u ON u.Id = la.UserId
        AND u.TenantId = la.TenantId
    WHERE (s.Date > DATEADD(d, - 90, GETDATE()) AND s.TenantId = @tenantId)
    GROUP BY COALESCE (u.Displayname, s.Address)
)
SELECT
	DisplayName, 
	Protection, 
	[Encryption], 
	Disclaimer, 
	LargeFiles, 
	FilesUploadedToSandbox,
	PdfMailsSent,
	SMimeMailsSigned,
	SMimeMailsEncrypted,
	SMimeMailsValidated,
	SMimeMailsDecrypted,
	PgpMailsSigned,
	PgpMailsEncrypted,
	PgpMailsValidated,
	PgpMailsDecrypted
FROM Query
WHERE (Protection <> 0) 
	OR ([Encryption] <> 0) 
	OR (Disclaimer <> 0) 
	OR (LargeFiles <> 0) 
	OR (FilesUploadedToSandbox <> 0)
