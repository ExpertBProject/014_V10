GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 13/11/2016 22:30:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(20), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE

-- VALIDACIONES ----------------------------------------------------------------------------------------------------------------
-- Todas
EXEC [dbo].[_98_EXO_CHK_OADM] @object_type, @transaction_type, @list_of_cols_val_tab_del, @error OUTPUT, @error_message OUTPUT
-- Sólo Sucursales y matriz
--EXEC [dbo].[_98_EXO_CHK_OCRD] @object_type, @transaction_type, @list_of_cols_val_tab_del, @error OUTPUT, @error_message OUTPUT
-- Sólo activo en las empresas de consolidación
EXEC [dbo].[_98_EXO_CHK_OICO] @object_type, @transaction_type, @list_of_cols_val_tab_del, @error OUTPUT, @error_message OUTPUT
--------------------------------------------------------------------------------------------------------------------------------

--------------------------------------------------------------------------------------------------------------------------------

-- Select the return values
select @error, @error_message

end