Set Dateformat DMY
DECLARE @FI as Datetime='01/04/2022'
DECLARE @FF as Datetime='30/04/2022'
DECLARE @TT AS varchar(2)='DS'

SELECT  SUM((Bascula.PesoLleno - Bascula.PesoVacio) / 1000 * VentasDetalle.Cantidad /100) AS Cantidad, PilasFisicas.Descripcion
FROM    PilasFisicas INNER JOIN
        Pilas ON PilasFisicas.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN
        VentasDetalle ON Pilas.IdPila = VentasDetalle.IdPila INNER JOIN
        Ventas ON VentasDetalle.IdVenta = Ventas.IdVentas INNER JOIN
        Bascula ON Ventas.IdVentas = Bascula.NumeroTransaccion
WHERE   (Bascula.TransaccionOrigen = @TT) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND YEAR(Bascula.FechaLleno)=2022
GROUP BY PilasFisicas.Descripcion
