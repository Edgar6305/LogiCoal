SET DATEFORMAT DMY

DECLARE @Fecha Datetime='03/07/2022 00:01:00' --GetDate()

DECLARE @Turno AS Varchar(1)
DECLARE @FechaTurno Date 
DECLARE @FechaINI Varchar(20)
DECLARE @FechaFIN Varchar(20)

/* Turno 1 de 06 a 18 */
IF DATEPART(HOUR, @Fecha)>=6 AND DATEPART(HOUR, @Fecha)<18
	Begin
		Set @Turno='1'
		Set @FechaTurno=CONVERT(date,@Fecha,103) 
		Set @FechaINI=CONVERT(varchar,@Fecha,103) + ' 06:00:00'
		Set @FechaFIN=CONVERT(varchar,@Fecha,103) + ' 18:00:00'

	End

/* Turno 2 de 18 a 23:59 */
IF DATEPART(HOUR, @Fecha)>=18 AND DATEPART(HOUR, @Fecha)<=23
	Begin
		Set @Turno='2'
		Set @FechaTurno=CONVERT(date,@Fecha,103) 
		Set @FechaINI=CONVERT(varchar,@Fecha,103) + ' 18:00:00'
		Set @FechaFIN=CONVERT(varchar,DATEADD(DAY,+1,@Fecha),103) + ' 06:00:00'
	End

/* Turno 2 de 24 a 05:59 */
IF DATEPART(HOUR, @Fecha)>=0 AND DATEPART(HOUR, @Fecha)<6
	Begin
		Set @Turno='2'
		Set @FechaTurno=CONVERT(date,DATEADD(DAY,-1,@Fecha),103) 
		Set @FechaINI=CONVERT(varchar,DATEADD(DAY,-1,@Fecha),103) + ' 18:00:00'
		Set @FechaFIN=CONVERT(varchar,@Fecha,103) + ' 06:00:00'
	End


Select @Turno, @FechaTurno, @FechaINI, @FechaFIN


