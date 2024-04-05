
/**************************************************************
				--- Cifras regionales----
***************************************************************/
  

 ---------------------------------------------------------------------------------
---------- Montos por Región ----
---------------------------------------------------------------------------------

--if OBJECT_ID('tempdb..#MONTOS_SECT') is not null drop table #MONTOS_SECT
	SELECT 		loc.region 
           , SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD'
		   , SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP'
		   ,count(oc.porid) CantOC

  -- into #MONTOS_SECT
    FROM  
    DM_Transaccional..THOrdenesCompra AS OC     
    inner JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC.IDFechaEnvioOC=TPO.DateKey 
    LEFT JOIN DM_Transaccional..DimProcedenciaOC AS PRO   ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
    left JOIN DM_Transaccional..DimComprador AS C        ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
    left JOIN DM_Transaccional..DimInstitucion AS I      ON C.entCode = I.entCode   
    LEFT JOIN DM_Transaccional..DimSector AS S           ON S.IdSector = I.IdSector
   left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

    WHERE   TPO.YEAR in (2023) 
    GROUP BY  Region
	order by Region,Monto_Bruto_USD desc

	

 ---------------------------------------------------------------------------------
---------- Montos por Región y tamaño empresa ----
---------------------------------------------------------------------------------

--if OBJECT_ID('tempdb..#MONTOS_SECT') is not null drop table #MONTOS_SECT
	SELECT 		loc.region 
			,  CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					ELSE 'MiPyme'
				    END				as			'Tmn'
           , SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD'
		   , SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP'
		   ,count(oc.porid) CantOC

  -- into #MONTOS_SECT
    FROM  
    DM_Transaccional..THOrdenesCompra AS OC     
    inner JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC.IDFechaEnvioOC=TPO.DateKey 
    LEFT JOIN DM_Transaccional..DimProcedenciaOC AS PRO   ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
    left JOIN DM_Transaccional..DimComprador AS C        ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
    left JOIN DM_Transaccional..DimInstitucion AS I      ON C.entCode = I.entCode   
    LEFT JOIN DM_Transaccional..DimSector AS S           ON S.IdSector = I.IdSector
   left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
   LEFT JOIN DM_Transaccional..DimProveedor AS PROV			ON PROV.IDSucursal=OC.IDSucursal
   LEFT JOIN DM_Transaccional..THTamanoProveedor AS TMN1		ON TMN1.entcode = PROV.entCode AND AñoTributario=(tpo.year -1) 

    WHERE   TPO.YEAR in (2023) 
    GROUP BY  Region,  CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					ELSE 'MiPyme'
				    END				
	order by Region,Monto_Bruto_USD desc



	   	 
---------------------------------------------------------------------------------
------------OCs con mayores montos por región 2023
---------------------------------------------------------------------------------
if OBJECT_ID('tempdb..#montoOCRegion') is not null drop table #montoOCRegion
		
SELECT 
oc.CodigoOC,
oc.NombreOC,
i.NombreInstitucion,
oc.link,
region, 
sum(MontoUSD + ImpuestoUSD) Monto_Bruto_USD,
sum(Montoclp + Impuestoclp) Monto_Bruto_CLP

into #montoOCRegion
	FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
	[DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
	[DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
	[DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
	[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
	left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
	where
	t.Year in (2023)
	group by region, oc.CodigoOC,oc.NombreOC,i.NombreInstitucion,oc.link
	order by Monto_Bruto_USD desc

Select *
FROM (Select *,
		Rank() over(partition by region order by Monto_Bruto_USD desc) as Rank
		FROM #montoOCRegion) as A
		where rank<=3
		order by region,rank 


---------------------------------------------------------------------------------
------------Rubros con mayores montos transados por región 2023
---------------------------------------------------------------------------------
if OBJECT_ID('tempdb..#montoRubrosRegion') is not null drop table #montoRubrosRegion
		
SELECT 
r.RubroN1,
region, 
sum(ocl.MontoUSD) Monto_USD,
sum(ocl.Montoclp) Monto_CLP

into #montoRubrosRegion
	FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
	 DM_Transaccional..THOrdenesCompralinea ocl on ocl.porID=oc.porID inner join
	[DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
	[DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
	[DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
	[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
	left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad	
	left join DM_Transaccional..DimProducto p on p.IDProducto=ocl.IDProducto
	left join DM_Transaccional..DimRubro r on r.IdRubro=p.IdRubro
	where
	t.Year in (2023)
	group by region, r.RubroN1
	order by region, sum(ocl.MontoUSD) desc


Select *
FROM (Select *,
		Rank() over(partition by region order by Monto_USD desc) as Rank
		FROM #montoRubrosRegion) as A
		where rank<=3
		order by region,rank 








/****************************************************************************************************************************************************************************/



---------------------------------------------------------------------------------
------------OOPP con mayores montos transados por región 2022
---------------------------------------------------------------------------------
if OBJECT_ID('tempdb..#montoOOPPRegion') is not null drop table #montoOOPPRegion
		
SELECT 
i.entCode,
NombreInstitucion Institución,
region, 
sum(MontoUSD + ImpuestoUSD) Monto_Bruto_USD,
sum(Montoclp + Impuestoclp) Monto_Bruto_CLP

into #montoOOPPRegion
	FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
	[DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
	[DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
	[DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
	[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
	left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
	where
	t.Year in (2022)
	group by region, i.entcode, NombreInstitucion
	order by sum(MontoUSD + ImpuestoUSD) desc, entCode

Select *
FROM (Select *,
		Rank() over(partition by region order by Monto_Bruto_USD desc) as Rank
		FROM #montoOOPPRegion) as A
		where rank<=3
		order by region,rank 

 

 
---------------------------------------------------------------------------------
------------Rubros con mayores montos transados por región 2022
---------------------------------------------------------------------------------
if OBJECT_ID('tempdb..#montoRubrosRegion') is not null drop table #montoRubrosRegion
		
SELECT 
r.RubroN1,
region, 
sum(ocl.MontoUSD) Monto_USD,
sum(ocl.Montoclp) Monto_CLP

into #montoRubrosRegion
	FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
	 DM_Transaccional..THOrdenesCompralinea ocl on ocl.porID=oc.porID inner join
	[DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
	[DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
	[DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
	[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
	left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad	
	left join DM_Transaccional..DimProducto p on p.IDProducto=ocl.IDProducto
	left join DM_Transaccional..DimRubro r on r.IdRubro=p.IdRubro
	where
	t.Year in (2022)
	group by region, r.RubroN1
	order by region, sum(ocl.MontoUSD) desc


Select *
FROM (Select *,
		Rank() over(partition by region order by Monto_USD desc) as Rank
		FROM #montoRubrosRegion) as A
		where rank<=3
		order by region,rank 




---------------------------------------------------------------------------------
------------OCs con mayores montos por región 2022
---------------------------------------------------------------------------------
if OBJECT_ID('tempdb..#montoOCRegion') is not null drop table #montoOCRegion
		
SELECT 
oc.CodigoOC,
oc.NombreOC,
i.NombreInstitucion,
oc.link,
region, 
sum(MontoUSD + ImpuestoUSD) Monto_Bruto_USD,
sum(Montoclp + Impuestoclp) Monto_Bruto_CLP

into #montoOCRegion
	FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
	[DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
	[DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
	[DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
	[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
	left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
	where
	t.Year in (2022)
	group by region, oc.CodigoOC,oc.NombreOC,i.NombreInstitucion,oc.link
	order by Monto_Bruto_USD desc

Select *
FROM (Select *,
		Rank() over(partition by region order by Monto_Bruto_USD desc) as Rank
		FROM #montoOCRegion) as A
		where rank<=3
		order by region,rank 



---------------------------------------------------------------------------------
------------Compra Ágil
---------------------------------------------------------------------------------

---------- Monto transado CAg en USD
SELECT	Region,
		SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD_CAg'	    ,
		SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP_CAg'	    ,
			count(distinct porid) CantOC_CAg
					FROM      DM_Transaccional..THOrdenesCompra AS OC     
					inner JOIN DM_Transaccional..DimTiempo AS TPO    ON OC.IDFechaEnvioOC=TPO.DateKey 
					inner join [DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra inner join
				[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
				inner JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
					   WHERE   TPO.YEAR in (2022) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
							and oc.porIsIntegrated=3
							group by Region


---------- Incremento real de monto transado CAg respecto al mismo periodo del año anterior e incremento en cantidad OC----
SELECT	loc.Region,
		SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD_CAg'	    ,
		SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP_CAg'	    ,
			count(distinct porid) CantOC_CAg,
			(select SUM(OC2.MontoUSD+OC2.ImpuestoUSD) 
			FROM      DM_Transaccional..THOrdenesCompra AS OC2     
					inner JOIN DM_Transaccional..DimTiempo AS TPO2    ON OC2.IDFechaEnvioOC=TPO2.DateKey 
					inner join DM_Transaccional.dbo.DimComprador c2 on c2.IDUnidaddeCompra=oc2.IDUnidaddeCompra inner join
					DM_Transaccional..diminstitucion i2 on i2.entcode=c2.entCode 
					inner JOIN DM_Transaccional..DimLocalidad as loc2	  ON C2.IDLocalidadUnidaddeCompra =  LOC2.IDLocalidad
			WHERE   TPO2.YEAR in (2021) and oc2.porIsIntegrated=3 and loc2.Region=loc.Region collate Modern_Spanish_CI_AS)  'MONTOUSD_CAg_2021',
			(select count(OC2.porid) 
			FROM      DM_Transaccional..THOrdenesCompra AS OC2     
					inner JOIN DM_Transaccional..DimTiempo AS TPO2    ON OC2.IDFechaEnvioOC=TPO2.DateKey 
					inner join DM_Transaccional.dbo.DimComprador c2 on c2.IDUnidaddeCompra=oc2.IDUnidaddeCompra inner join
					DM_Transaccional..diminstitucion i2 on i2.entcode=c2.entCode 
					inner JOIN DM_Transaccional..DimLocalidad as loc2	  ON C2.IDLocalidadUnidaddeCompra =  LOC2.IDLocalidad
			WHERE   TPO2.YEAR in (2021) and oc2.porIsIntegrated=3 and loc2.Region=loc.Region collate Modern_Spanish_CI_AS)  'CantOC_CAg_2021'

					FROM      DM_Transaccional..THOrdenesCompra AS OC     
					inner JOIN DM_Transaccional..DimTiempo AS TPO    ON OC.IDFechaEnvioOC=TPO.DateKey 
					inner join [DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra inner join
				[DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
				inner JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
					   WHERE   TPO.YEAR in (2022) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
							and oc.porIsIntegrated=3
							group by loc.Region
							order by loc.Region


