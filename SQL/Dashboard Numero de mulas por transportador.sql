SET DATEFORMAT DMY
SELECT   Transportador.Descripcion, Count(*) Numero
FROM     Bascula INNER JOIN
         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador
WHERE    (Bascula.FechaTurno = CONVERT(DATETIME, '2022-10-05 00:00:00', 102)) AND (Bascula.IdTransaccion = 2) AND Bascula.Estado='AC'
GROUP BY Transportador.Descripcion

