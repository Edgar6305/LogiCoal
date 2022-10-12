Set dateformat DMY
DECLARE @FI DateTime = '01/08/2022 06:00'
DECLARE @FF DateTime = '31/08/2022 18:00'

--SELECT   Descripcion,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=1) AS R1,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=2) AS R2,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=3) AS R3,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=4) AS R4,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=5) AS R5,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=6) AS R6,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=7) AS R7,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=8) AS R8,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=9) AS R9,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=10) AS R10,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=11) AS R11,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=12) AS R12,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=13) AS R13,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=14) AS R14,
--(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote WHERE (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=TP.IdTipoCarbon AND  DAY(FechaTurno)=15) AS R15

--FROM      TiposCarbon AS TP
--WHERE     IdTipoCarbon=1

--SELECT   Descripcion,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=1  AND TipoCarbon=TP.IdTipoCarbon) AS PR1,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=2  AND TipoCarbon=TP.IdTipoCarbon) AS PR2,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=3  AND TipoCarbon=TP.IdTipoCarbon) AS PR3,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=4  AND TipoCarbon=TP.IdTipoCarbon) AS PR4,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=5  AND TipoCarbon=TP.IdTipoCarbon) AS PR5,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=6  AND TipoCarbon=TP.IdTipoCarbon) AS PR6,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=7  AND TipoCarbon=TP.IdTipoCarbon) AS PR7,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=8  AND TipoCarbon=TP.IdTipoCarbon) AS PR8,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=9  AND TipoCarbon=TP.IdTipoCarbon) AS PR9,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=10 AND TipoCarbon=TP.IdTipoCarbon) AS PR10,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=11 AND TipoCarbon=TP.IdTipoCarbon) AS PR11,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=12 AND TipoCarbon=TP.IdTipoCarbon) AS PR12,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=13 AND TipoCarbon=TP.IdTipoCarbon) AS PR13,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=14 AND TipoCarbon=TP.IdTipoCarbon) AS PR14,
--(SELECT  Cantidad AS Presupuesto FROM PlanesMinerosDetalle WHERE Dia=15 AND TipoCarbon=TP.IdTipoCarbon) AS PR15

--FROM     TiposCarbon AS TP
--WHERE    IdTipoCarbon=1


--SELECT   DAY(Bascula.FechaTurno) AS Dia, Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad 
--FROM     Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote 
--WHERE    (Bascula.TransaccionOrigen = 'LT') AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND Lotes.IdTipoCarbon=1  
--Group By DAY(Bascula.FechaTurno)
--Order By DAY(Bascula.FechaTurno)



Select Dia,
	(SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad 
	FROM Bascula INNER JOIN	 Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote 
	WHERE (Bascula.TransaccionOrigen = 'LT') AND Lotes.IdTipoCarbon=1 AND (Bascula.FechaTurno >= @FI) AND (Bascula.FechaTurno <= @FF ) AND DAY(FechaTurno)=Dias.dia) AS Cantidad,

	(SELECT SUM(Cantidad) AS Presupuesto FROM PlanesMinerosDetalle WHERE TipoCarbon=1 AND Mes=Month(@FI) AND anio=Year(@FI) AND Dia=Dias.Dia) AS Presupuesto

From Dias 

