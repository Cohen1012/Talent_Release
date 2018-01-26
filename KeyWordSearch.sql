USE [TalentDatabase]
GO

/****** Object:  StoredProcedure [dbo].[KeyWordSearch]    Script Date: 2017/12/21 上午 12:54:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[KeyWordSearch] 
	-- Add the parameters for the stored procedure here 
	@KeyWord Nvarchar(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	DECLARE @SearchStr nvarchar(200) = @KeyWord


-- Copyright © 2002 Narayana Vyas Kondreddi. All rights reserved.
-- Purpose: To search all columns of all tables for a given search string
-- Written by: Narayana Vyas Kondreddi
-- Site: http://vyaskn.tripod.com
-- Tested on: SQL Server 7.0 and SQL Server 2000
-- Date modified: 28th July 2002 22:50 GMT


CREATE TABLE #Results (ColumnName nvarchar(370), ColumnValue nvarchar(3630))

SET NOCOUNT ON

DECLARE @TableName nvarchar(256), @ColumnName nvarchar(128), @SearchStr2 nvarchar(110)
SET  @TableName = ''
SET @SearchStr2 = QUOTENAME('%' + @SearchStr + '%','''')

WHILE @TableName IS NOT NULL
BEGIN
	SET @ColumnName = ''
	SET @TableName = 
	(
		SELECT MIN(QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME))
		FROM 	INFORMATION_SCHEMA.TABLES
		WHERE 		TABLE_TYPE = 'BASE TABLE'
			AND	QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME) > @TableName
			AND	OBJECTPROPERTY(
					OBJECT_ID(
						QUOTENAME(TABLE_SCHEMA) + '.' + QUOTENAME(TABLE_NAME)
						 ), 'IsMSShipped'
					       ) = 0
	)

	WHILE (@TableName IS NOT NULL) AND (@ColumnName IS NOT NULL)
	BEGIN
		SET @ColumnName =
		(
			SELECT MIN(QUOTENAME(COLUMN_NAME))
			FROM 	INFORMATION_SCHEMA.COLUMNS
			WHERE 		TABLE_SCHEMA	= PARSENAME(@TableName, 2)
				AND	TABLE_NAME	= PARSENAME(@TableName, 1)
				AND	DATA_TYPE IN ('char', 'varchar', 'nchar', 'nvarchar')
				AND	QUOTENAME(COLUMN_NAME) > @ColumnName
		)

		IF @ColumnName IS NOT NULL
		BEGIN
			INSERT INTO #Results
			EXEC
			(
				'SELECT ''' + @TableName + '.' + @ColumnName + ''', LEFT(' + @ColumnName + ', 3630) 
				FROM ' + @TableName + ' (NOLOCK) ' +
				' WHERE ' + @ColumnName + ' LIKE ' + @SearchStr2
			)
		END
	END	
END

SELECT * FROM #Results

DROP TABLE #Results
END
GO