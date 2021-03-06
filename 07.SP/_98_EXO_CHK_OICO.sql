
GO
/****** Object:  StoredProcedure [dbo].[_98_EXO_CHK_OICO]    Script Date: 12/12/2016 16:04:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[_98_EXO_CHK_OICO]
	-- Add the parameters for the stored procedure here
	@object_type NVARCHAR(20),
	@transaction_type NCHAR(1),
	@list_of_cols_val_tab_del NVARCHAR(255),
	@error INT OUTPUT,
	@error_message NVARCHAR(200) OUTPUT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    --Variables
	DECLARE @U_EXO_DBNAME NVARCHAR(100)
	DECLARE @CONT INT
	DECLARE @LineId INT
	DECLARE @U_EXO_DBNAME1 NVARCHAR(100)
	DECLARE @PRCNTVALIDO NVARCHAR(1)
	DECLARE @Sucursal NVARCHAR(100)
	
	IF @object_type = 'EXO_OICO'
		BEGIN
			IF @transaction_type IN ('A', 'U')
				BEGIN
					SET @U_EXO_DBNAME = ''

					SELECT @U_EXO_DBNAME = ISNULL(t1.U_EXO_DBNAME, '')
					FROM [@EXO_OICO] t1 WITH (NOLOCK)
					WHERE t1.DocEntry = @list_of_cols_val_tab_del

					SET @CONT = ISNULL((SELECT COUNT(t1.DocEntry)
									   FROM [@EXO_ICO1] t1 WITH (NOLOCK)
									   WHERE t1.DocEntry = @list_of_cols_val_tab_del), 0)

					SET @Sucursal = ISNULL((SELECT ISNULL(t1.U_EXO_DBNAME, '') U_EXO_DBNAME
											FROM [@EXO_ICO1] t1 WITH (NOLOCK)
											WHERE t1.DocEntry = @list_of_cols_val_tab_del
											GROUP BY ISNULL(t1.U_EXO_DBNAME, '')
											HAVING COUNT(t1.DocEntry) > 1), '')

					SELECT @LineId = t1.LineId, @U_EXO_DBNAME1 = ISNULL(t1.U_EXO_DBNAME, ''), 
					@PRCNTVALIDO = CASE WHEN ISNULL(t1.U_EXO_PRCNT, 0) < CAST(0 AS NUMERIC(19, 6)) OR ISNULL(t1.U_EXO_PRCNT, 0) > CAST(100 AS NUMERIC(19, 6)) THEN 'N' ELSE 'Y' END
					FROM [@EXO_ICO1] t1 WITH (NOLOCK)
					WHERE t1.DocEntry = @list_of_cols_val_tab_del
					AND (ISNULL(t1.U_EXO_DBNAME, '') = '' OR 
					CASE WHEN ISNULL(t1.U_EXO_PRCNT, 0) <= CAST(0 AS NUMERIC(19, 6)) OR ISNULL(t1.U_EXO_PRCNT, 0) > CAST(100 AS NUMERIC(19, 6)) THEN 'N' ELSE 'Y' END = 'N')

					IF @U_EXO_DBNAME = ''
						BEGIN
							Set @error = 1
							Set @error_message = '(EXO) El campo Empresa consolidación es obligatorio.'
						END
					ELSE
						IF @CONT = 0
							BEGIN
								Set @error = 1
								Set @error_message = '(EXO) Debe añadir al menos una Sucursal.'
							END
						ELSE
							IF @U_EXO_DBNAME1 = ''
								BEGIN
									Set @error = 1
									Set @error_message = '(EXO) El campo Sucuarsal es obligatorio.'
								END
							ELSE
								IF @PRCNTVALIDO = 'N'
									BEGIN
										Set @error = 1
										Set @error_message = '(EXO) El campo % consolidación debe de ser mayor o igual que 0 y menor o igual que 100.'
									END
								ELSE
									IF @Sucursal <> ''
										BEGIN
											Set @error = 1
											Set @error_message = '(EXO) La Sucursal ' + @Sucursal + ' ya existe.'
										END
				END
		END
END

