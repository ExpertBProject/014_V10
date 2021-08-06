SELECT Tabla.ClavePed, Tabla.DocNum, Tabla.CuentaProvi, Tabla.CuentaGasto, Tabla.CeCo1, Tabla.CeCo2, Tabla.CeCo3, Tabla.Projecto, sum(Tabla.TotProvisionar)AS 'TotalProvisionar' FROM
(
  SELECT TCab.DocEntry AS 'ClavePed', TCab.DocNum as 'DocNum', (1 - TCab.DiscPrcnt / 100.0 ) * TLin.Price * TLin.OpenCreQty AS 'TotProvisionar',
           TLin.AcctCode AS 'CuentaGasto',
           dbo.EXO_CuentProvArt('##AVANZADA', TCab.FinncPriod, TLin.ItemCode, TLin.WhsCode) AS 'CuentaProvi',           
           isnull(TLin.OcrCode, '') AS 'CeCo1', isnull(TLin.OcrCode2, '') AS 'CeCo2', isnull(TLin.OcrCode3, '') AS 'CeCo3',    
           TLin.Project AS 'Projecto'
     FROM OPOR TCab 
     	INNER JOIN  POR1 TLin ON TLin.DocEntry = TCab.DocEntry
	   	INNER JOIN OITM TArt ON TLin.ItemCode = TArt.ItemCode
      	WHERE  TCab.DocEntry = ##CLAVEPEDIDO
      	AND  TLin.LineStatus = 'O' AND TArt.InvntItem = 'N'      	      	
      	
) AS Tabla
GROUP BY Tabla.ClavePed, Tabla.DocNum, Tabla.CuentaProvi, Tabla.CuentaGasto, Tabla.CeCo1, Tabla.CeCo2, Tabla.CeCo3, Tabla.Projecto