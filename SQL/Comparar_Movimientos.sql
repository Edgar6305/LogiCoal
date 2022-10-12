SET DATEFORMAT DMY
DECLARE @FechaIni as Datetime='02/05/2022'

SELECT  Bascula.IdTiquete, Bascula.PesoLleno - Bascula.PesoVacio AS PesoNeto, Terceros.Identificacion AS Tercero, Bascula.FechaLleno
FROM    Bascula INNER JOIN
        Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN
        Terceros ON Ventas.IdCliente = Terceros.IdCliente
WHERE   (Bascula.TransaccionOrigen = 'DS') AND (Bascula.Estado = 'AC') 

SELECT  Bascula.IdTiquete, Bascula.PesoLleno-Bascula.PesoVacio AS Pesoneto, 
		CASE WHEN Acopios.Ubicacion='TOLU' THEN '409'
		     WHEN Acopios.Ubicacion='RIVERPORT' THEN '410'  END AS Tercero,  Bascula.FechaLleno
FROM   Pilas INNER JOIN
       Acopios ON Pilas.IdAcopio = Acopios.IdAcopio INNER JOIN
       Bascula INNER JOIN
       Traslados ON Bascula.NumeroTransaccion = Traslados.IdTraslado ON Pilas.IdPila = Traslados.PilaDestino
WHERE  Bascula.TransaccionOrigen='TR' AND Bascula.Estado='AC' 