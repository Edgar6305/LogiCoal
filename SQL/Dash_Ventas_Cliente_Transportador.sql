DECLARE @TRA AS Varchar(MAX)='1,3,6,10,15,16,20,21,25,28'
DECLARE @CLI AS Varchar(2)='20'

SELECT  Transportador.Descripcion,

(SELECT ISNULL(ROUND(SUM(Bascula.PesoLleno - Bascula.PesoVacio) / 1000, 0), 0) 
FROM    Bascula INNER JOIN
        Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN
        Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN
        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador
WHERE  (Bascula.TransaccionOrigen = 'DS') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaVacio) = 2022) AND (MONTH(Bascula.FechaVacio) = 5) 
		AND Terceros.IdCliente=@CLI AND Transportador.IdTransportador=TR.value) AS Neto

FROM String_Split(@TRA,',') AS TR inner join Transportador ON Transportador.IdTransportador=TR.value




