Set Dateformat DMY
DECLARE @FI as Datetime='01/05/2022'
DECLARE @FF as Datetime='30/05/2022'

SELECT ROUND(SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000,0) AS Neto, PilasFisicas.Descripcion
FROM     Bascula INNER JOIN
                  Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote INNER JOIN
                  Pilas ON Lotes.Pila = Pilas.IdPila INNER JOIN
                  PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica
WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado <> 'AN') AND (Bascula.IdMaterial = 1) AND (Bascula.FechaLleno >=@FI ) AND (Bascula.FechaLleno <=@FF )
GROUP BY PilasFisicas.Descripcion