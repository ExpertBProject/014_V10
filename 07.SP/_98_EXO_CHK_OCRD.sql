GO
/****** Object:  StoredProcedure [dbo].[_98_EXO_CHK_OCRD]    Script Date: 13/11/2016 21:06:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[_98_EXO_CHK_OCRD]
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
	DECLARE @U_EXO_GRUPOEMPRESA NVARCHAR(10)
	DECLARE @CardType NVARCHAR(1)
	DECLARE @CONT INT
	
	IF @object_type = '2'
		BEGIN
			IF @transaction_type IN ('A', 'U')
				BEGIN
					SELECT @U_EXO_GRUPOEMPRESA = ISNULL(t1.U_EXO_GRUPOEMPRESA, '-'), @CardType = t1.CardType
					FROM OCRD t1 WITH (NOLOCK)
					WHERE t1.CardCode = @list_of_cols_val_tab_del
										
					SET @CONT = ISNULL((SELECT COUNT(t1.CardCode)
					                    FROM OCRD t1 WITH (NOLOCK)
										WHERE t1.CardCode <> @list_of_cols_val_tab_del
										AND ISNULL(t1.U_EXO_GRUPOEMPRESA, '-') = @U_EXO_GRUPOEMPRESA
										AND t1.CardType = @CardType), 0)
					
					IF @CONT <> 0 AND @U_EXO_GRUPOEMPRESA <> '-'
						BEGIN
							Set @error = 1
							Set @error_message = '(EXO) Ya existe un IC de tipo ' + @CardType + ' con el Grupo de empresas indicado.'
						END
				END
		END
END
