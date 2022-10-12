SELECT *
FROM OPENROWSET('Microsoft.ACE.OLEDB.16.0','Excel 12.0;HDR=NO;Database=C:\FRONTIER\Movimientos\2022\TOLU\Report Frontier_tolu_21_MV_BAY_PEARL.xlsx','Select F2,F16,F17,F20 From [Basico_Tractomula$] WHERE ISDATE(F16) AND F2>0');

--USE [master];
--EXEC sys.sp_configure 'show advanced options', 1;
--RECONFIGURE;
--EXEC sys.sp_configure 'Ad Hoc Distributed Queries', 1;
--RECONFIGURE;

--USE [master] 
--GO 
--EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'AllowInProcess', 1 
--GO 
--EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.16.0', N'DynamicParameters', 1 
--GO 

--DECLARE @sql char(150)
--DECLARE @Driver sysname
--DECLARE @Cadena sysname
--DECLARE @DirName sysname ='C:\FRONTIER\Movimientos\2022\TOLU\Report Frontier_tolu_21_MV_BAY_PEARL.xlsx'
--DECLARE @Sentencia sysname

--SET @Driver= ''''+'Microsoft.ACE.OLEDB.16.0'+''''
--SET @Cadena=''''+'Excel 12.0;HDR=NO;Database=' + @DirName +''''
--SET @Sentencia=''''+'Select * from [Hoja1$]'+''''

--SET @sql='SELECT * FROM OPENROWSET('+@Driver+', '+@Cadena+', '+@Sentencia+')'
--Select @sql
--execute (@sql)