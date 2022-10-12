Set Dateformat DMY
DECLARE @FI1 as Datetime='11/04/2022 06:00:00'
DECLARE @FF1 as Datetime='11/04/2022 18:00:00'

DECLARE @FI2 as Datetime='11/04/2022 18:00:00'
DECLARE @FF2 as Datetime='11/05/2022 06:00:00'
DECLARE @TT AS varchar(2)='LT'

SELECT  SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000) AS Cantidad,
	(SELECT  SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000)
	FROM    Bascula
	WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND MONTH(Bascula.FechaLleno)=4) AS Mes,

	(SELECT  SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000)
	FROM    Bascula
	WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaVacio>=@FI1) AND 
			(Bascula.FechaVacio<=@FF2)) AS Dia,

	(SELECT  ISNULL(SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000),0)
	FROM    Bascula
	WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaVacio>=@FI1) AND 
			(Bascula.FechaVacio<=@FF1)) AS Turno1,

	(SELECT  ISNULL(SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000),0)
	FROM    Bascula
	WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaVacio>=@FI2) AND 
			(Bascula.FechaVacio<=@FF2)) AS Turno2

FROM    Bascula
WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND YEAR(Bascula.Fechavacio)=2022
