USE [ServiceDeskEmailReminder]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--Uncomment function if function is not existing
--IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_GetClosureDate') AND type in (N'F'))
--BEGIN
--CREATE FUNCTION [dbo].[FN_GetClosureDate]
--(
--	@pdteDate DATE
--)
--RETURNS DATE
--AS
--BEGIN
--	--add one day each day
--	-- check each day from start and end date (if holiday or weekend is found, plus 1)
--	WHILE (DATEPART(dw, @pdteDate) IN (1,7) OR EXISTS (SELECT NULL FROM tbl_Holidays WHERE cast(DateofHoliday as date) = @pdteDate ))
--	BEGIN
--		SET @pdteDate = DATEADD(DAY,1,@pdteDate)
--	END
--	RETURN @pdteDate
--END
--END

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[DateValidation]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[DateValidation](
	[validateDate] [nvarchar](20) NULL,
	[validateForOpen] [nvarchar](100) NULL,
	[validateForSLA] [nvarchar](100) NULL,
	[validateFname] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImportExcel]') AND type in (N'U'))
BEGIN 
CREATE TABLE [dbo].[ImportExcel](
	[importID] [int] IDENTITY(1,1) NOT NULL,
	[importRequestID] [nvarchar](100) NULL,
	[importTitle] [nvarchar](100) NULL,
	[importOwner] [nvarchar](100) NULL,
	[importCurrentStatus] [nvarchar](100) NULL,
	[importSubmittedTime] [nvarchar](100) NULL,
	[importSubmittedBy] [nvarchar](100) NULL,
	[importWorkgroup] [nvarchar](100) NULL,
	[importCategory] [nvarchar](100) NULL,
	[importSubCategory] [nvarchar](100) NULL,
	[importSubSubCategory] [nvarchar](100) NULL,
	[importLastModifiedTime] [nvarchar](100) NULL,
	[importSLAStatus] [nvarchar](100) NULL,
	[importResolutionTime] [nvarchar](100) NULL,
	[importExpectedClosureTime] [nvarchar](100) NULL,
	[importExpectedSLAClosureTime] [nvarchar](100) NULL,
	[importStateCategory] [nvarchar](100) NULL,
	[importDepartment] [nvarchar](100) NULL,
	[importDepartment1] [nvarchar](100) NULL,
	[importSource] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[NoOwner]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[NoOwner](
	[noID] [int] IDENTITY(1,1) NOT NULL,
	[noTicketTitle] [nvarchar](100) NULL,
	[noStaffName] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ProcessorDetails]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[ProcessorDetails](
	[processorID] [int] IDENTITY(1,1) NOT NULL,
	[processorStaffName] [nvarchar](100) NULL,
	[processorStaffNickname] [nvarchar](100) NULL,
	[processorStaffEmail] [nvarchar](100) NULL,
	[processorForSLACCEmail] [nvarchar](100) NULL,
	--[processorForOpenEmail]  [nvarchar](100) NULL,
	[processorMandatoryCCEmail] [nvarchar](100) NULL,
	[processorToMeEmail] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO


IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbl_Holidays]') AND type in (N'U'))
BEGIN
Create table [dbo].[tbl_Holidays](
[HolidayID] [int] IDENTITY(1,1) NOT NULL,
[Holidayname][nvarchar](50)  NOT NULL,
[DateofHoliday][Date] NULL
--[isRegular][bit] null
)
END


IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[USP_Holiday]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[USP_Holiday] AS' 
END
GO
ALTER PROCEDURE [dbo].[USP_Holiday] 
(
@pintmode as INT,
@pHolidayid as INT,
@pstrHolidayname AS NVARCHAR(100),
@pdteDateofHoliday AS DATE
--@pisRegular AS BIT
)
AS 
--Selecting Holidays in tbl_Holidays
IF @pintMode=1
BEGIN	
	SELECT Holidayid [Holiday ID],
		   Holidayname [Holiday Name],
		   DateofHoliday [Date of Holiday]
  FROM tbl_Holidays

END
-- when updating Holiday tbl
ELSE IF @pintmode=2
  BEGIN UPDATE tbl_Holidays
	SET Holidayname	  =	@pstrHolidayname,
		DateofHoliday = @pdteDateofHoliday
	    --isRegular     = @pisRegular
	WHERE Holidayid=@pHolidayid
END

--Delete from tbl_Holiday
ELSE IF @pintmode = 3
  BEGIN DELETE FROM tbl_Holidays 
		WHERE Holidayid=@pHolidayid
		END


--Adding New Holidays in tbl_Holidays 
ELSE IF @pintmode = 4
	INSERT INTO tbl_Holidays(Holidayname,
							 DateofHoliday)
	VALUES (@pstrHolidayname,
			@pdteDateofHoliday)

 

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[USP_CRUD]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[USP_CRUD] AS' 
END
GO
ALTER PROCEDURE [dbo].[USP_CRUD]
(
@pintMode AS INT,
@pintID AS INT,
@pstrStaffName AS NVARCHAR(100),
@pstrStaffNickname AS NVARCHAR(100),
@pstrStaffEmail AS NVARCHAR(100),
@pstrForSLACCEmail AS NVARCHAR(100),  
--@pstrForOpenEmail AS NVARCHAR(100),
@pstrTicketTitle AS NVARCHAR(100),
@pstrNoStaffName AS NVARCHAR(100)
)
AS

DECLARE @pstrMandatoryCCEmail  NVARCHAR(100),
		@pstrToMeEmail  NVARCHAR(100)
SET @pstrMandatoryCCEmail = 'cyrene.alcala@filinvestland.com, anthony.bondoc@filinvestland.com'
SET @pstrToMeEmail = 'cyrene.alcala@filinvestland.com'
-- Calling Processor Details Table for DGV
IF (@pintMode=1)
BEGIN
	SELECT processorID,
		   processorStaffName [Processor Name],
		   processorStaffNickname [Processor Nickname],
		   processorStaffEmail [Email Address],
		   processorForSLACCEmail [For SLA CC Email]
		   --processorForOpenEmail [For Open CC Email]
	FROM ProcessorDetails
END
             
-- Calling No Owner Table for DGV
ELSE IF (@pintMode=2)
BEGIN
	SELECT noID,
		   noStaffName [Processor Name],
		   noTicketTitle [Ticket Title]
	FROM NoOwner
END

-- Adding data in Processor Details Table
IF (@pintMode=3)
BEGIN
	INSERT INTO ProcessorDetails (processorStaffName,
									  processorStaffNickname,
									  processorStaffEmail,
									  processorForSLACCEmail,
									  --processorForOpenEmail,
									  processorMandatoryCCEmail,
									  processorToMeEmail)
	VALUES (@pstrStaffName,
			@pstrStaffNickname,
			@pstrStaffEmail,
			@pstrForSLACCEmail,
			--@pstrForOpenEmail,
			@pstrMandatoryCCEmail,
			@pstrToMeEmail)
END

-- Adding in No Owner Table
ELSE IF (@pintMode=4)
BEGIN
	INSERT INTO NoOwner (noStaffName,
							 noTicketTitle)
	VALUES (@pstrNoStaffName,
			@pstrTicketTitle)
END

-- triggers when updating Processor Details Table
IF (@pintMode=5)
BEGIN
	UPDATE ProcessorDetails
	SET processorStaffName=@pstrStaffName,
		processorStaffNickname=@pstrStaffNickname,
		processorStaffEmail=@pstrStaffEmail,
		processorForSLACCEmail=@pstrForSLACCEmail
		--processorForOpenEmail = @pstrForOpenEmail
	WHERE processorID=@pintID
END

-- Updating in No Owner Table
ELSE IF (@pintMode=6)
BEGIN
	UPDATE NoOwner
	SET noStaffName=@pstrNoStaffName,
		noTicketTitle=@pstrTicketTitle
	WHERE noID=@pintID
END

-- Deletion in Processor Details Table
IF (@pintMode=7)
BEGIN
	DELETE FROM ProcessorDetails
	WHERE processorID=@pintID
END

-- Deletion in No Owner Table
IF (@pintMode=8)
BEGIN
	DELETE FROM NoOwner
	WHERE noID=@pintID
END

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[USP_EmailSendingProcess]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[USP_EmailSendingProcess] AS' 
END
GO
ALTER PROCEDURE [dbo].[USP_EmailSendingProcess]
(
@pintMode INT,
@pbitIsToProcessor BIT,
@pstrDateNow NVARCHAR(20),
@pstrVOpen CHAR(2),
@pstrFname NVARCHAR(max)
)

AS
-- Creation of Main Temporary Table 
-- be used to store imported excel file details
CREATE TABLE #MainTempTable
(
mttRequestID NVARCHAR(100),
mttTitle NVARCHAR(100),
mttOwner NVARCHAR(100),
mttCurrentStatus NVARCHAR(100),
mttSubmittedTime NVARCHAR(100),
mttSubmittedBy NVARCHAR(100),
mttWorkgroup NVARCHAR(100),
mttCategory NVARCHAR(100),
mttSubCategory NVARCHAR(100),
mttSubSubCategory NVARCHAR(100),
mttLastModifiedTime NVARCHAR(100),
mttSLAStatus NVARCHAR(100),
mttResolutionTime NVARCHAR(100),
mttExpectedClosureTime NVARCHAR(100),
mttExpectedSLAClosureTime NVARCHAR(100),
mttStateCategory NVARCHAR(100),
mttDepartment NVARCHAR(100),
mttDepartment1 NVARCHAR(100),
mttSource NVARCHAR(100)
)
--temporary table for Open Tickets
CREATE TABLE #TempTableForOpenTickets
(
fotRequestID NVARCHAR(100),
fotOwner NVARCHAR(100),
fotCurrentStatus NVARCHAR(100),
fotSubmittedTime NVARCHAR(100),
fotCategory NVARCHAR(100),
fotSLAStatus NVARCHAR(100),
fotExpectedSLAClosureTime NVARCHAR(100)
)
--temporary table for SLA Tickets
CREATE TABLE #TempTableForSLATickets
(
fstRequestID NVARCHAR(100),
fstOwner NVARCHAR(100),
fstCurrentStatus NVARCHAR(100),
fstSubmittedTime NVARCHAR(100),
fstCategory NVARCHAR(100),
fstSLAStatus NVARCHAR(100),
fstExpectedSLAClosureTime NVARCHAR(100)
)

CREATE TABLE #TempTableOwnerOnly
(
ootOwner NVARCHAR(100)
)

CREATE TABLE #TempTableWithNicknameForEmail
(
ttnStaffName NVARCHAR(100),
ttnStaffEmail NVARCHAR(100),
ttnStaffNickname NVARCHAR(100),
ttnSLACCEmail NVARCHAR(100),
--ttnOpenCCEmail NVARCHAR (100),
ttnMandatoryCCEmail NVARCHAR(100)
)


									/*THE PROCESS*/

-- Inserting the data of imported excel file from ImportExcel(table) to MainTempTable
INSERT INTO #MainTempTable (mttRequestID,
							mttTitle,
							mttOwner,
							mttCurrentStatus,
							mttSubmittedTime,
							mttSubmittedBy,
							mttWorkgroup,
							mttCategory,
							mttSubCategory,
							mttSubSubCategory,
							mttLastModifiedTime,
							mttSLAStatus,
							mttResolutionTime,
							mttExpectedClosureTime,
							mttExpectedSLAClosureTime,
							mttStateCategory,
							mttDepartment,
							mttDepartment1,
							mttSource)
SELECT importRequestId,
	   importTitle,
CASE
	WHEN importOwner = '' THEN noStaffName
	WHEN importOwner = NULL then noStaffName
	ELSE importOwner
END

AS	importOwner,
	importCurrentStatus,
	importSubmittedTime,
	importSubmittedBy,
	importWorkgroup,
	importCategory,
	importSubCategory,
	importSubSubCategory,
	importLastModifiedTime,
	importSLAStatus,
	importResolutionTime,
	importExpectedClosureTime,
	importExpectedSLAClosureTime,
	importStateCategory,
	importDepartment,
	importDepartment1,    
	importSource
FROM ImportExcel
LEFT JOIN NoOwner
ON importTitle=noTicketTitle

-- Sorting of OPEN Tickets from MainTempTable(including Program and Process Enhancements) 
IF (@pintMode=1)
	BEGIN
	INSERT INTO #TempTableForOpenTickets (fotRequestID,
										  fotOwner,
										  fotCurrentStatus,
										  fotSubmittedTime,
										  fotCategory,
										  fotSLAStatus,
										  fotExpectedSLAClosureTime)
	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime

	FROM #MainTempTable
	WHERE
		mttCurrentStatus NOT LIKE 'Promoted to Production'
		AND mttCurrentStatus NOT LIKE '%Data%'
		AND mttCurrentStatus NOT LIKE '%Resolved%'
		AND mttCurrentStatus NOT LIKE 'Cancel'
		AND mttOwner != ''
	ORDER BY mttRequestID

	SELECT  fotRequestID [Request ID],
			fotOwner [Owner],
			fotCurrentStatus [Status],
			fotSubmittedTime [Submitted Time],
			fotCategory [Category],
			fotSLAStatus [SLA Status],
			fotExpectedSLAClosureTime [Expected SLA Closure Time]
	FROM #TempTableForOpenTickets
END

-- Sorting of OPEN Tickets from MainTempTable(NOT Including Program and Process Enhancements)
IF (@pintMode=2)
BEGIN
	INSERT INTO #TempTableForOpenTickets (fotRequestID,
										  fotOwner,
										  fotCurrentStatus,
										  fotSubmittedTime,
										  fotCategory,
										  fotSLAStatus,
										  fotExpectedSLAClosureTime)
	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime
	FROM #MainTempTable
	WHERE
		mttCurrentStatus NOT LIKE 'Promoted to Production(Resolved)'
		AND mttCurrentStatus NOT LIKE 'Data%'
		AND mttCurrentStatus NOT LIKE 'Resolved%'
		AND mttCurrentStatus NOT LIKE 'Cancel'
		AND mttCategory NOT LIKE 'Program and Process Enhancement%'
		AND mttOwner != ''
	ORDER BY mttRequestID

	SELECT  fotRequestID [Request ID],
			fotOwner [Owner],
			fotCurrentStatus [Status],
			fotSubmittedTime [Submitted Time],
			fotCategory [Category],
			fotSLAStatus [SLA Status],
			fotExpectedSLAClosureTime [Expected SLA Closure Time]
	FROM #TempTableForOpenTickets
END

-- Sorting of SLA (INCLUDING Program and Process Enhancements) Tickets from MainTempTable
-- Current Date and Greater than TIME tickets
IF (@pintMode=3)
BEGIN
	INSERT INTO #TempTableForSLATickets (fstRequestID,
									 fstOwner,
									 fstCurrentStatus,
									 fstSubmittedTime,
									 fstCategory,
									 fstSLAStatus,
									 fstExpectedSLAClosureTime)
	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime
	FROM #MainTempTable
	WHERE
		NOT EXISTS (SELECT validateDate
					FROM DateValidation
					WHERE mttExpectedSLAClosureTime=validateDate)
		AND mttExpectedSLAClosureTime != 'On Hold'
		AND mttExpectedSLAClosureTime != '--'
		AND mttExpectedSLAClosureTime != 'N/A'
		AND mttOwner != ''
		AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = FORMAT(CAST(GETDATE()AS DATE), 'MM/dd/yyyy')
		AND CONVERT(TIME(0),mttExpectedSLAClosureTime) > CONVERT(TIME(0), GETDATE())
	ORDER BY mttExpectedSLAClosureTime

--OPEN
-- Current Date of imported tickets > Current Date + 1

DECLARE	@mdteClosureDate	DATE = dbo.FN_GetClosureDate(DATEADD(DAY,1,GETDATE()))
	INSERT INTO #TempTableForSLATickets (fstRequestID,
									 fstOwner,
									 fstCurrentStatus,
									 fstSubmittedTime,
									 fstCategory,
									 fstSLAStatus,
									 fstExpectedSLAClosureTime)
	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime
	FROM #MainTempTable
	LEFT JOIN tbl_Holidays 
					ON MONTH(mttExpectedSLAClosureTime) = MONTH(DateofHoliday) AND
					DAY(mttExpectedSLAClosureTime) = MONTH(DateofHoliday)
	WHERE
		NOT EXISTS (SELECT validateDate
					FROM DateValidation
					WHERE mttExpectedSLAClosureTime=validateDate)
		AND mttExpectedSLAClosureTime != 'On Hold'
		AND mttExpectedSLAClosureTime != '--'
		AND mttExpectedSLAClosureTime != 'N/A'
		AND mttOwner != ''
		AND (FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = @mdteClosureDate or CAST(mttExpectedSLAClosureTime AS DATE)=CAST(GETDATE() AS DATE))
		--AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = @mdteClosureDate																	
		--AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = FORMAT(CAST(GETDATE()+1 AS DATE), 'MM/dd/yyyy')
		--AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = CASE DATEPART(DW, GETDATE()+1)
																				--WHEN 6 THEN FORMAT(CAST(GETDATE()+3 AS DATE), 'MM/dd/yyyy') --Friday
																				--WHEN 7 THEN FORMAT(CAST(GETDATE()+2 AS DATE), 'MM/dd/yyyy') --Saturday
																				--ELSE FORMAT(CAST(GETDATE()+1 AS DATE), 'MM/dd/yyyy')
																		--	END
				
	ORDER BY mttExpectedSLAClosureTime

	SELECT DISTINCT fstRequestID [Request ID],
					fstOwner [Owner],
					fstCurrentStatus [Status],
					fstSubmittedTime [Submitted Time],
					fstCategory [Category],
					fstSLAStatus [SLA Status],
					fstExpectedSLAClosureTime [Expected SLA Closure Time]
	FROM #TempTableForSLATickets
END

-- Sorting of SLA (NOT INCLUDING Program and Process Enhancements) Tickets from MainTempTable
-- Current Date and Greater than TIME tickets (NO PAPE)
ELSE IF (@pintMode=4)
BEGIN
	INSERT INTO #TempTableForSLATickets (fstRequestID,
									 fstOwner,
									 fstCurrentStatus,
									 fstSubmittedTime,
									 fstCategory,
									 fstSLAStatus,
									 fstExpectedSLAClosureTime)

	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime
	FROM #MainTempTable
	WHERE
		NOT EXISTS (SELECT mttExpectedSLAClosureTime
					FROM DateValidation
					WHERE mttExpectedSLAClosureTime=validateDate)
		AND mttExpectedSLAClosureTime != 'On Hold'
		AND mttExpectedSLAClosureTime != '--'
		AND mttExpectedSLAClosureTime != 'N/A'
		AND mttCategory != 'Program and Process Enhancement'
		AND mttOwner != ''
		AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = FORMAT(CAST(GETDATE()AS DATE), 'MM/dd/yyyy')
		AND CONVERT(TIME(0),mttExpectedSLAClosureTime) > CONVERT(TIME(0), GETDATE())
	ORDER BY mttExpectedSLAClosureTime

--DECLARE	@mdteClosureDate	DATE = DATEADD(DAY,1,GETDATE())
--WHILE (DATEPART(dw, @mdteClosureDate) IN (1,7) OR EXISTS (SELECT NULL FROM tbl_Holidays WHERE cast(DateofHoliday as date) = @mdteClosureDate ))
--BEGIN
--	SET @mdteClosureDate = DATEADD(DAY,1,@mdteClosureDate)
--END



-- SLA
-- Current Date of imported tickets > Current Date + 1 (NO PAPE) 
DECLARE	@mdteeClosureDate	DATE = dbo.FN_GetClosureDate(DATEADD(DAY,1,GETDATE()))
INSERT INTO #TempTableForSLATickets (fstRequestID,
									 fstOwner,
									 fstCurrentStatus,
									 fstSubmittedTime,
									 fstCategory,
									 fstSLAStatus,
									 fstExpectedSLAClosureTime)
 
	SELECT mttRequestID,
		   mttOwner,
		   mttCurrentStatus,
		   mttSubmittedTime,
		   mttCategory,
		   mttSLAStatus,
		   mttExpectedSLAClosureTime
	FROM #MainTempTable
	LEFT JOIN tbl_Holidays
			ON MONTH(mttExpectedSLAClosureTime) = MONTH(DateofHoliday) AND
			   DAY(mttExpectedSLAClosureTime) = DAY(DateofHoliday)
	WHERE
		NOT EXISTS (SELECT validateDate
					FROM DateValidation
					WHERE mttExpectedSLAClosureTime=validateDate)
		AND mttExpectedSLAClosureTime != 'On Hold'
		AND mttExpectedSLAClosureTime != '--'
		AND mttExpectedSLAClosureTime != 'N/A'
		AND mttCategory != 'Program and Process Enhancement'
		AND mttOwner != ''
		AND (FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = @mdteeClosureDate or CAST(mttExpectedSLAClosureTime AS DATE)=CAST(GETDATE() AS DATE))
		--AND FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = @mdteeClosureDate
		--AND	FORMAT(CAST(mttExpectedSLAClosureTime AS DATE), 'MM/dd/yyyy') = FORMAT(CAST(GETDATE()+1 AS DATE), 'MM/dd/yyyy')
	ORDER BY mttExpectedSLAClosureTime

	SELECT DISTINCT fstRequestID [Request ID],
					fstOwner [Owner],
					fstCurrentStatus [Status],
					fstSubmittedTime [Submitted Time],
					fstCategory [Category],
					fstSLAStatus [SLA Status],
					fstExpectedSLAClosureTime [Expected SLA Closure Time]
	FROM #TempTableForSLATickets
	ORDER BY fstExpectedSLAClosureTime
END

/* END OF PROCESS */
----------------------------------------------------------------------------------------------------------------------
	
-- Email Sending (With Condition) for OPEN Tickets 
 IF (@pintMode=5)
BEGIN
	INSERT INTO #TempTableOwnerOnly(ootOwner)
	SELECT DISTINCT mttOwner
	FROM #MainTempTable
	WHERE mttOwner != ''
	ORDER BY mttOwner

	
	INSERT INTO #TempTableWithNicknameForEmail (ttnStaffName,
												ttnStaffEmail,
												ttnStaffNickname,
												--ttnOpenCCEmail,
												ttnMandatoryCCEmail)
	SELECT	ootOwner, 
			CASE @pbitIsToProcessor 
				WHEN 0 THEN processorToMeEmail
				WHEN 1 THEN processorStaffEmail 
			END ,
			processorStaffNickname,
			processorMandatoryCCEmail
			--processorForOpenEmail
	FROM #TempTableOwnerOnly
	LEFT JOIN ProcessorDetails
	ON ootOwner=processorStaffName 
	WHERE ootOwner != '' 
		--AND ootOwner != 'Mary Ann C. Tan' 
		--AND ootOwner != 'Rolando dela Cruz'
		--AND ootOwner != 'Grace Marie M. Bada' 
	ORDER BY ootOwner

	SELECT * FROM #TempTableWithNicknameForEmail
END

-- Email Sending (With Condition) for SLA Tickets
ELSE IF (@pintMode=6)
BEGIN
	INSERT INTO #TempTableOwnerOnly(ootOwner)
	SELECT DISTINCT mttOwner
	FROM #MainTempTable
	WHERE mttOwner != ''
	ORDER BY mttOwner

	
	INSERT INTO #TempTableWithNicknameForEmail(ttnStaffName,
												ttnStaffEmail,
												ttnStaffNickname,
												ttnSLACCEmail,
												ttnMandatoryCCEmail)
	SELECT	ootOwner, 
			CASE @pbitIsToProcessor 
				WHEN 0 THEN processorToMeEmail
				WHEN 1 THEN processorStaffEmail 
			END ,
			processorStaffNickname,
			processorForSLACCEmail, 
			processorMandatoryCCEmail
	FROM #TempTableOwnerOnly
	LEFT JOIN ProcessorDetails
	ON ootOwner=processorStaffName
	WHERE ootOwner != '' 
		--AND ootOwner != 'Mary Ann C. Tan' 
		--AND ootOwner != 'Rolando dela Cruz'
		--AND ootOwner != 'Grace Marie M. Bada'
	ORDER BY ootOwner

	SELECT * FROM #TempTableWithNicknameForEmail
	
END
-- once sending only
-- Validation if User already sent and Email Reminder for SLA Tickets
IF (@pintMode=7)
BEGIN
	IF EXISTS (SELECT validateDate
			   FROM DateValidation
			   WHERE validateDate = @pstrDateNow)
	BEGIN
		UPDATE DateValidation
		SET validateForSLA = @pstrVOpen,
			validateFname = @pstrFname
		WHERE validateDate = @pstrDateNow
	END
ELSE
	BEGIN
		INSERT INTO DateValidation (validateDate,
										validateFname)
		VALUES (@pstrDateNow,
				@pstrFname)
		UPDATE DateValidation
		SET validateForSLA = @pstrVOpen,
			validateFname = @pstrFname
		WHERE validateDate = @pstrDateNow
	END
END
-- once sending only
-- Validation if User already sent and Email Reminder for OPEN Tickets 
ELSE IF (@pintMode=8)
BEGIN
	IF EXISTS (SELECT validateDate
			   FROM DateValidation
			   WHERE validateDate = @pstrDateNow)
	BEGIN
		UPDATE DateValidation
		SET validateForOpen = @pstrVOpen,
			validateFname = @pstrFname
		WHERE validateDate = @pstrDateNow
	END
ELSE
	BEGIN
		INSERT INTO DateValidation (validateDate,
										validateFname)
		VALUES (@pstrDateNow,
				@pstrFname)
		UPDATE DateValidation
		SET validateForOpen = @pstrVOpen,
			validateFname = @pstrFname
		WHERE validateDate = @pstrDateNow
	END
END

-- Selecting DateValidation table
IF (@pintMode=9)
BEGIN
	SELECT * FROM DateValidation
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[USP_ExcelToDBProcess]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[USP_ExcelToDBProcess] AS' 
END
GO
ALTER PROCEDURE [dbo].[USP_ExcelToDBProcess]
(
@pintMode INT,
@pstrRequestID NVARCHAR(100),
@pstrTitle NVARCHAR(100),
@pstrOwner NVARCHAR(100),
@pstrCurrentStatus NVARCHAR(100),
@pstrSubmittedtime NVARCHAR(100),
@pstrSubmittedBy NVARCHAR(100),
@pstrWorkgroup NVARCHAR(100),
@pstrCategory NVARCHAR(100),
@pstrSubCategory NVARCHAR(100),
@pstrSubSubCategory NVARCHAR(100),
@pstrLastModifiedTime NVARCHAR(100),
@pstrSLAStatus NVARCHAR(100),
@pstrResolutionTime NVARCHAR(100),
@pstrExpectedClosuretime NVARCHAR(100),
@pstrExpectedSLAClosureTime NVARCHAR(100),
@pstrStateCategory NVARCHAR(100),
@pstrDepartment NVARCHAR(100),
@pstrDepartment1 NVARCHAR(100),
@pstrSource NVARCHAR(100)
)

AS

IF @pintMode=1
BEGIN
	INSERT INTO ImportExcel(importRequestID,
						   importTitle,
						   importOwner,
						   importCurrentStatus,
						   importSubmittedTime,
						   importSubmittedBy,
						   importWorkgroup,
						   importCategory,
						   importSubCategory,
						   importSubSubCategory,
						   importLastModifiedTime,
						   importSLAStatus,
						   importResolutionTime,
						   importExpectedClosureTime,
						   importExpectedSLAClosureTime,
						   importStateCategory,
						   importDepartment,
						   importDepartment1,
						   importSource)
	VALUES (@pstrRequestID,
			@pstrTitle,
			@pstrOwner,
			@pstrCurrentStatus,
			@pstrSubmittedtime,
			@pstrSubmittedBy,
			@pstrWorkgroup,
			@pstrCategory,
			@pstrSubCategory,
			@pstrSubSubCategory,
			@pstrLastModifiedTime,
			@pstrSLAStatus,
			@pstrResolutionTime,
			@pstrExpectedClosuretime,
			@pstrExpectedSLAClosureTime,
			@pstrStateCategory,
			@pstrDepartment,
			@pstrDepartment1,
			@pstrSource)
END

ELSE IF @pintMode=2

BEGIN
	DELETE FROM ImportExcel
END

GO

