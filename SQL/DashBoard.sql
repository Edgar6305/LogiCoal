Set Dateformat DMY
DECLARE @FI1 as Datetime='11/04/2022 06:00:00'
DECLARE @FF1 as Datetime='11/04/2022 18:00:00'

DECLARE @FI2 as Datetime='11/04/2022 18:00:00'
DECLARE @FF2 as Datetime='11/05/2022 06:00:00'

SELECT SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000 AS Anio,
	(SELECT SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000
	FROM     Bascula 
	WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND MONTH(Bascula.FechaLleno)=4) AS Mes,

	(SELECT SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000
	FROM     Bascula 
	WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaLleno>=@FI1) AND 
	(Bascula.FechaLleno<=@FF2)) AS Dia,

	(SELECT SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000
	FROM     Bascula 
	WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaLleno>=@FI1) AND 
	(Bascula.FechaLleno<=@FF1)) AS Turno1,

	(SELECT SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000
	FROM     Bascula 
	WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaLleno>=@FI2) AND 
	(Bascula.FechaLleno<=@FF2)) AS Turno2



FROM     Bascula 
WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND Year(Bascula.FechaLleno)=2022
