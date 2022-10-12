SET DateFormat DMY
DECLARE @Turno int
DECLARE @FechaTurno Date
DECLARE @Fecha Datetime=Getdate()

EXEC PA_FechaTurno_Turno @Fecha, @Turno OUTPUT
EXEC PA_FechaTurno_FechaTurno @Fecha, @FechaTurno OUTPUT

SELECT    ISNULL(COUNT(*), 0) AS Viajes, OperadoresMineros.Descripcion
FROM      Bascula INNER JOIN
          Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote INNER JOIN
          OperadoresMineros ON Lotes.Operador = OperadoresMineros.IdOperador
WHERE    (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado = 'AC') AND (Bascula.FechaTurno = @FechaTurno)
GROUP BY Lotes.Operador, OperadoresMineros.Descripcion




