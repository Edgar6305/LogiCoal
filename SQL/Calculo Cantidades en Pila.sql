DECLARE @Pila  INT=3
DECLARE @Saldo Float=0

--> Lotes Suma a Pila Origen(ROM) 
Select @Saldo=ISNULL(Sum(Cantidad),0) From Lotes Where Pila=@Pila

--> Trituracion Resta a Pila Origen(ROM) 
Select @Saldo=@Saldo-ISNULL(-Sum(Cantidad),0) From Trituracion Where PilaOrigen=@Pila

--> Trituracion Suma A Pila Destino
SELECT @Saldo=@Saldo + ISNULL(Sum(Trituracion.Cantidad*TrituracionDetalle.Porcentaje/100),0) 
FROM   Trituracion INNER JOIN TrituracionDetalle ON Trituracion.IdTrituracion = TrituracionDetalle.IdTrituracion
WHERE  (TrituracionDetalle.PilaDestino = @Pila)

--> Traslados Suma A Pila Destino los Traslados
SELECT @Saldo=@Saldo + ISNULL(Sum(Traslados.CantidadDespachada),0)
FROM   Traslados INNER JOIN TrasladosDetalle ON Traslados.IdTraslado = TrasladosDetalle.IdTraslado
WHERE  Traslados.PilaDestino=@Pila

--> Traslados Resta De Pila Origen los Traslados
SELECT @Saldo=@Saldo - ISNULL(Sum(Traslados.CantidadDespachada*TrasladosDetalle.Porcentaje/100),0)
FROM   Traslados INNER JOIN TrasladosDetalle ON Traslados.IdTraslado = TrasladosDetalle.IdTraslado
WHERE  TrasladosDetalle.IdPilaOrigen = @Pila

--> Ventas Resta DE Pilas Origen
SELECT @Saldo=@Saldo - ISNULL(Sum(Ventas.CantidadDespachada*VentasDetalle.Cantidad),0)
FROM   Ventas INNER JOIN VentasDetalle ON Ventas.IdVentas = VentasDetalle.IdVenta
WHERE  VentasDetalle.IdPila=@Pila

Select @Saldo



