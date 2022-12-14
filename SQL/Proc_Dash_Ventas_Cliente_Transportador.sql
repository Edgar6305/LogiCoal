USE [TRACER]
GO
/****** Object:  UserDefinedFunction [dbo].[FT_DB_RecepcionHora]    Script Date: 25/05/2022 14:07:06 p.m. ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE FUNCTION  [dbo].[FT_DB_VentasPorTransportador] (@CLI INT, @TRA AS Varchar(MAX), @ANIO as INT, @MES as INT) 

RETURNS TABLE 

RETURN (
SELECT  Transportador.Descripcion,

(SELECT ISNULL(ROUND(SUM(Bascula.PesoLleno - Bascula.PesoVacio) / 1000, 0), 0) 
FROM    Bascula INNER JOIN
        Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN
        Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN
        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador
WHERE  (Bascula.TransaccionOrigen = 'DS') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaVacio) = @ANIO) AND (MONTH(Bascula.FechaVacio) = @MES) 
		AND Terceros.IdCliente=@CLI AND Transportador.IdTransportador=TR.value) AS Neto

FROM String_Split(@TRA,',') AS TR inner join Transportador ON Transportador.IdTransportador=TR.value
)

