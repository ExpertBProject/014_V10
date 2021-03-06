
GO
/****** Object:  StoredProcedure [dbo].[_98_EXO_CHK_OADM]    Script Date: 10/11/2016 17:23:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[_98_EXO_CHK_OADM]
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
	DECLARE @U_EXO_CONSOLIDACION NVARCHAR(1)
	DECLARE @U_EXO_MATRIZ NVARCHAR(1)
	DECLARE @U_EXO_GRUPOEMPRESA NVARCHAR(10)
	
	IF @object_type = '39'
		BEGIN
			IF @transaction_type IN ('U')
				BEGIN
					SELECT @U_EXO_CONSOLIDACION = ISNULL(t1.U_EXO_CONSOLIDACION, 'N'), @U_EXO_MATRIZ = ISNULL(t1.U_EXO_MATRIZ, 'N'),
					@U_EXO_GRUPOEMPRESA = ISNULL(t1.U_EXO_GRUPOEMPRESA, '')
					FROM OADM t1 WITH (NOLOCK)
					
					IF @U_EXO_CONSOLIDACION = 'Y' AND @U_EXO_MATRIZ = 'Y'
						BEGIN
							Set @error = 1
							Set @error_message = '(EXO) La empresa no puede ser Consolidación y Matriz a la vez. Seleccione una opción u otra.'
						END
					ELSE
						IF @U_EXO_CONSOLIDACION = 'Y' AND @U_EXO_GRUPOEMPRESA <> ''
							BEGIN
								Set @error = 1
								Set @error_message = '(EXO) Las Empresas de consolidación no tienen que tener Grupo de empresas.'
							END
						ELSE
							IF @U_EXO_CONSOLIDACION = 'N' AND @U_EXO_GRUPOEMPRESA = ''
								BEGIN
									Set @error = 1
									Set @error_message = '(EXO) El campo Grupo de empresas es obligatorio.'
								END
				END
		END
END
