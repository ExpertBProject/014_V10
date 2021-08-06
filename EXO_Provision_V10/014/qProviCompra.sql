DECLARE @HASTAFECHA AS DATETIME

SELECT @HASTAFECHA = CONVERT(DATETIME, '##HASTAFECHA', 112)

SELECT T0.DocEntry AS 'Clave', T0.DocNum AS 'NumDoc', T0.TaxDate AS 'Fecha Doc', T0.NumAtCard AS 'Ref Prov',
	   T0.CardCode AS 'Proveedor', T0.CardName AS 'Nombre', 
	   T0.DocTotal AS 'Total Doc', 
	   ( SELECT (1 - T0.DiscPrcnt / 100.0 ) * SUM(TLin.Price * TLin.OpenCreQty )   FROM POR1 TLin 
	   		  INNER JOIN OITM TArt ON TLin.ItemCode = TArt.ItemCode
      		  WHERE TLin.DocEntry = T0.DocEntry AND  TLin.LineStatus = 'O' AND TArt.InvntItem = 'N') AS 'Total Provisionar'	   	   	   	   
 FROM OPOR T0
WHERE T0.U_EXO_ProviComp = 'Y' AND ISNULL(T0.U_EXO_AsiProvComp, 0) = 0 AND T0.TaxDate <= @HASTAFECHA AND
      EXISTS (SELECT 'TRUE' FROM POR1 TLin INNER JOIN OITM TArt ON TLin.ItemCode = TArt.ItemCode
      		  WHERE TLin.DocEntry = T0.DocEntry AND  TLin.LineStatus = 'O' AND TArt.InvntItem = 'N')
ORDER BY T0.DocNum
      

