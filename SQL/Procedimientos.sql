Select dbo.FS_CantidadLotes (1, 'LT', 2, 'AC')  

EXEC PA_PesoInicial 1,1,'LT',2,'R1',1,1,'UGQ367', 'Edgardo Hernandez' , 1,28555	,'sisma','AC'

EXEC PA_PesoFinal 2,3,12822 

EXEC PA_CierreLotes 2,'sisma'

EXEC PA_CreaRegCalidad 'LT',2,2,12,'sisma'
