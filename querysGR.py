#################################
###   Querys para generador   ###
#################################


#----------------------------------#
# Query con datos a nivel regional #
#----------------------------------#
def QueryReg():
    q = '''
        ;with temp as (
        SELECT TPO.YEAR 'Ano'
            , tpo.Month as 'Mes'
            , LOC.Region
        , (CASE oc1.porisintegrated WHEN 3 THEN 'CAg'
                    else (case  OC1.IDProcedenciaOC
                                    WHEN 703 THEN 'Convenio Marco'
                                    WHEN 701 THEN 'Licitación Pública'
                                    WHEN 1401 THEN 'Licitación Pública'
                                    WHEN 702 THEN 'Licitación Privada'
                                    ELSE 'Trato Directo' END)END) 'ProcedenciaOC'
        , I.NombreInstitucion
        ,case when a.idTamano is not null then cast(a.idTamano as nvarchar)
                    else isnull(b.tamanonombre,'SinDato') end tamano0
        ,  COUNT(DISTINCT P.RUTSucursal)      'CantProveedores'
        , COUNT(DISTINCT OC1.CodigoOC) 'CantOC'    
        , SUM(OC1.MontoUSD+OC1.ImpuestoUSD) 'MONTOUSD'
        , SUM(OC1.MontoCLP+OC1.ImpuestoCLP) 'MONTOCLP'
        , SUM(OC1.MontoCLF+OC1.ImpuestoCLF) 'MONTOCLF'



        FROM DM_Transaccional..THOrdenesCompra  AS OC1
        inner JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC1.IDFechaEnvioOC=TPO.DateKey
        inner JOIN DM_Transaccional..DimProcedenciaOC AS PRO   ON PRO.IDProcedenciaOC = OC1.IDProcedenciaOC
        left JOIN DM_Transaccional..DimComprador AS C        ON OC1.IDUnidaddeCompra = C.IDUnidaddeCompra
        left JOIN DM_Transaccional..DimProveedor AS P        ON P.IDSucursal=OC1.IDSucursal
        left join [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and AñoTributario=2021
        left join Estudios.dbo.TamanoProveedorNuevos20230802 b on p.entcode=b.entCode
        LEFT JOIN  DM_Transaccional..DimTamanoProveedor AS TP ON TP.IdTamano=a.IdTamano
        left JOIN DM_Transaccional..DimInstitucion AS I      ON C.entCode = I.entCode
        left JOIN DM_Transaccional..DimSector AS S           ON S.IdSector = I.IdSector
        left JOIN DM_Transaccional..DimLocalidad as loc      ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

        --Dado que existe comparacion anual, debe tomarse con año 1 y año -1
        WHERE   TPO.YEAR in(2023,2022)
                AND TPO.MONTH >= 1 AND TPO.MONTH <= 8
        GROUP BY TPO.YEAR,TPO.MONTH, 
        case when a.idTamano is not null then cast(a.idTamano as nvarchar)
                    else isnull(b.tamanonombre,'SinDato') end ,
                    I.NombreInstitucion,  LOC.Region,
            (CASE OC1.porisintegrated WHEN 3 THEN 'CAg'
                    else (case  OC1.IDProcedenciaOC
                                    WHEN 703 THEN 'Convenio Marco'
                                    WHEN 701 THEN 'Licitación Pública'
                                    WHEN 1401 THEN 'Licitación Pública'
                                    WHEN 702 THEN 'Licitación Privada'
                                    ELSE 'Trato Directo' END)END)
        )
        Select 		Ano

            , Mes
            , Region
        ,  ProcedenciaOC
        , NombreInstitucion
        ,case  when tamano0 in ('PYME','Microempresa','2','3','4') then 'MiPyme'
                    when tamano0 in ('Grande','1') then 'Grande'
                    when tamano0  in ('5','Extranjera','SinDato') then 'SinDato' 
                        end Tamano,
                sum(MONTOUSD) MONTOUSD,
                sum(MONTOCLF) MONTOCLF,
                sum(MONTOCLP) MONTOCLP,
                sum(CantOC) CantOC,
                sum(CantProveedores)      'CantProveedores'
        from temp
        group by 	Ano   
            , Mes
            , Region
        ,  ProcedenciaOC
        , NombreInstitucion
        , case  when tamano0 in ('PYME','Microempresa','2','3','4') then 'MiPyme'
                    when tamano0 in ('Grande','1') then 'Grande'
                    when tamano0  in ('5','Extranjera','SinDato') then 'SinDato' 
                        end
        '''
    return q


#--------------------------------------------------------#
# Query con datos a nivel nacional. Solo no tiene Region #
#--------------------------------------------------------#
def QueryTotal():
    q = '''
        ;with temp as (
            SELECT TPO.YEAR 'Ano' --TPO.MONTHNAME 'Mes'
            , tpo.Month as 'Mes'
            , (CASE oc1.porisintegrated WHEN 3 THEN 'CAg'
                    else (case  OC1.IDProcedenciaOC
                                    WHEN 703 THEN 'Convenio Marco'
                                        WHEN 701 THEN 'Licitación Pública'
                                        WHEN 1401 THEN 'Licitación Pública'
                                        WHEN 702 THEN 'Licitación Privada'
                                        ELSE 'Trato Directo' END)END) 'ProcedenciaOC'
            ,I.NombreInstitucion
            ,case when a.idTamano is not null then cast(a.idTamano as nvarchar)
                            else isnull(b.tamanonombre,'SinDato') end tamano0
            , COUNT(DISTINCT P.RUTSucursal)      'CantProveedores'
            , COUNT(DISTINCT OC1.CodigoOC) 'CantOC'    
            , SUM(OC1.MontoUSD+OC1.ImpuestoUSD) 'MONTOUSD'
            , SUM(OC1.MontoCLP+OC1.ImpuestoCLP) 'MONTOCLP'
            , SUM(OC1.MontoCLF+OC1.ImpuestoCLF) 'MONTOCLF'

            FROM DM_Transaccional..THOrdenesCompra  AS OC1
            inner JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC1.IDFechaEnvioOC=TPO.DateKey
            inner JOIN DM_Transaccional..DimProcedenciaOC AS PRO   ON PRO.IDProcedenciaOC = OC1.IDProcedenciaOC
            left JOIN DM_Transaccional..DimComprador AS C        ON OC1.IDUnidaddeCompra = C.IDUnidaddeCompra
            left JOIN DM_Transaccional..DimProveedor AS P        ON P.IDSucursal=OC1.IDSucursal
            left join [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and AñoTributario=2021
            left join Estudios.dbo.TamanoProveedorNuevos20230802 b on p.entcode=b.entCode
            LEFT JOIN  DM_Transaccional..DimTamanoProveedor AS TP ON TP.IdTamano=a.IdTamano
            left JOIN DM_Transaccional..DimInstitucion AS I      ON C.entCode = I.entCode
            left JOIN DM_Transaccional..DimSector AS S           ON S.IdSector = I.IdSector
            left JOIN DM_Transaccional..DimLocalidad as loc      ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
            
            --Dado que existe comparacion anual, debe tomarse con año 1 y año -1
            WHERE   TPO.YEAR in(2023,2022)
                    AND TPO.MONTH >= 1 AND TPO.MONTH <= 8
            GROUP BY TPO.YEAR,TPO.MONTH, TP.Tamano,
            case when a.idTamano is not null then cast(a.idTamano as nvarchar)
                        else isnull(b.tamanonombre,'SinDato') end ,
                        I.NombreInstitucion,
                (CASE oc1.porisintegrated WHEN 3 THEN 'CAg'
                        else (case  OC1.IDProcedenciaOC
                                        WHEN 703 THEN 'Convenio Marco'
                                        WHEN 701 THEN 'Licitación Pública'
                                        WHEN 1401 THEN 'Licitación Pública'
                                        WHEN 702 THEN 'Licitación Privada'
                                        ELSE 'Trato Directo' END)END)
        )
        Select       Ano

                , Mes
            ,  ProcedenciaOC
            , NombreInstitucion
            ,case  when tamano0 in ('PYME','Microempresa','2','3','4') then 'MiPyme'
                        when tamano0 in ('Grande','1') then 'Grande'
                        when tamano0  in ('5','Extranjera','SinDato') then 'SinDato'
                            end Tamano,
                    sum(MONTOUSD) MONTOUSD,
                    sum(MONTOCLF) MONTOCLF,
                    sum(MONTOCLP) MONTOCLP,
                    sum(CantOC) CantOC,
                    sum(CantProveedores)      'CantProveedores'
            from temp
            group by    Ano
                , Mes
            ,  ProcedenciaOC
            , NombreInstitucion
            , case  when tamano0 in ('PYME','Microempresa','2','3','4') then 'MiPyme'
                        when tamano0 in ('Grande','1') then 'Grande'
                        when tamano0  in ('5','Extranjera','SinDato') then 'SinDato'
                            end
        '''
    return q


#-------------------------------------------------#
# Montos totales transados por región 2022 y 2023 #
#-------------------------------------------------#
def queryTotalRegion():
    q = '''
        SELECT		TPO.YEAR 'Año'
                    ,loc.region 'Region'
                    , SUM(OC.MontoUSD+OC.ImpuestoUSD)/1000000 'Monto_Bruto_USD'
                    , SUM(OC.MontoCLP+OC.ImpuestoCLP)/1000000 'Monto_Bruto_CLP'
                    ,count(oc.porid) CantOC
                
        FROM  
            DM_Transaccional..THOrdenesCompra AS OC     
            inner JOIN DM_Transaccional..DimTiempo AS TPO           ON OC.IDFechaEnvioOC=TPO.DateKey 
            LEFT JOIN DM_Transaccional..DimProcedenciaOC AS PRO		ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
            left JOIN DM_Transaccional..DimComprador AS C			ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
            left JOIN DM_Transaccional..DimInstitucion AS I			ON C.entCode = I.entCode   
            LEFT JOIN DM_Transaccional..DimSector AS S				ON S.IdSector = I.IdSector
            left JOIN DM_Transaccional..DimLocalidad as loc			ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

        WHERE   TPO.YEAR in (2023)
            AND		TPO.MONTH>=1
            AND     TPO.MONTH<=8
        GROUP BY  Region,
                    TPO.YEAR
            --SECTOR

        UNION

        SELECT		TPO.YEAR 'Año'
                    ,loc.region collate Modern_Spanish_CI_AS 'Region'
                    ,SUM(OC.MontoUSD+OC.ImpuestoUSD)/1000000 'Monto_Bruto_USD'
                    ,SUM(OC.MontoCLP+OC.ImpuestoCLP)/1000000 'Monto_Bruto_CLP'
                    ,count(oc.porid) CantOC
                
        FROM  
            [10.34.71.227].DM_Transaccional_2022.DBO.THOrdenesCompra AS OC     
            inner JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimTiempo AS TPO			ON OC.IDFechaEnvioOC=TPO.DateKey 
            LEFT JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimProcedenciaOC AS PRO		ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
            left JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimComprador AS C			ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
            left JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimInstitucion AS I			ON C.entCode = I.entCode   
            LEFT JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimSector AS S				ON S.IdSector = I.IdSector
            left JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimLocalidad as loc			ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

        WHERE   TPO.YEAR in (2022)
            AND		TPO.MONTH>=1
            AND     TPO.MONTH<=8
        GROUP BY  Region
                ,TPO.YEAR

        order by Region
                ,TPO.YEAR
        '''
    return q


#------------------------------------------------#
# Montos totales transados por sector por región #
#------------------------------------------------#
def querySectorRegion():
    q = '''
        SELECT 	loc.region,
            Sector,
            SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD',
            SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP',
            count(oc.porid) CantOC
            

        FROM DM_Transaccional..THOrdenesCompra AS OC     
            inner JOIN DM_Transaccional..DimTiempo AS TPO			ON OC.IDFechaEnvioOC=TPO.DateKey
            left JOIN DM_Transaccional..DimComprador AS C			ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
            left JOIN DM_Transaccional..DimInstitucion AS I			ON C.entCode = I.entCode   
            LEFT JOIN DM_Transaccional..DimSector AS S				ON S.IdSector = I.IdSector
            left JOIN DM_Transaccional..DimLocalidad as loc			ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

        WHERE   TPO.YEAR in (2023) 
            AND	TPO.MONTH<=8
        GROUP BY  Region,
            SECTOR
        order by Region,
            Monto_Bruto_USD DESC, 
            SECTOR
        '''
    return q


#-----------------------------------------------#
# Montos totales transados por rubro por región #
#-----------------------------------------------#
def queryRubroRegion():
    q = ''' 
        select * 
        from( select
        LOC.Region
        ,rubro.RubroN1
        ,SUM(OCL.MontoUSD) 'MONTOUSD'
        , SUM(OCL.MontoCLP) 'MONTOCLP'
        , SUM(OCL.MontoCLF) 'MONTOCLF'
        , Rank() over (Partition BY LOC.Region ORDER BY  SUM(OCL.MontoUSD) DESC ) AS Rank
            FROM  
            DM_Transaccional..THOrdenesCompra AS OC     
            LEFT JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC.IDFechaEnvioOC=TPO.DateKey 
            LEFT JOIN DM_Transaccional..DimProcedenciaOC AS PRO   ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
            INNER JOIN DM_Transaccional..DimComprador AS C        ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra
            INNER JOIN DM_Transaccional..DimProveedor AS P        ON P.IDSucursal=OC.IDSucursal
            --LEFT JOIN  DM_Transaccional..DimTamanoProveedor AS TP ON TP.IdTamano=P.IdTamano
            INNER JOIN DM_Transaccional..DimInstitucion AS I      ON C.entCode = I.entCode   
            INNER JOIN DM_Transaccional..DimSector AS S           ON S.IdSector = I.IdSector
            INNER JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
            LEFT JOIN  DM_Transaccional..THOrdenesCompraLinea OCL ON OCL.porID=OC.porID
            LEFT JOIN DM_Transaccional..DimProducto produc ON produc.IDProducto=OCL.IDProducto 
            LEFT JOIN DM_Transaccional..DimRubro rubro on rubro.IdRubro=produc.IdRubro
            --Dado que existe comparacion anual, debe tomarse con año 1 y año -1
            WHERE   TPO.YEAR = 2023
                AND  TPO.month between 1 and 8
            GROUP BY LOC.Region,rubro.RubroN1
                    )rs WHERE Rank <= 3 ORDER BY Region ASC, Rank ASC 
        '''
    return q


#-----------------------------------------#
# Query CA por región periodo 2023 y 2022 #
#-----------------------------------------#
def queryCompraAgilRegion():
    q = '''
        DECLARE @ANO Int
        DECLARE @ANOM1 Int
        SET @ANO = 2023
        SET @ANOM1 = @ANO - 1

        SELECT	@ANO 'Ano',
                loc.Region 'Region',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP_CAg'	    ,
                SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD_CAg'	    ,
                count(distinct porid) 'CantOC_CAg'
        FROM	dm_transaccional..THOrdenesCompra AS OC     
                inner JOIN dm_transaccional..DimTiempo AS TPO		ON OC.IDFechaEnvioOC=TPO.DateKey 
                inner join [dm_transaccional].dbo.DimComprador c	on c.IDUnidaddeCompra=oc.IDUnidaddeCompra
                inner join [dm_transaccional]..diminstitucion i		on i.entcode=c.entCode 
                inner JOIN dm_transaccional..DimLocalidad as loc	ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
        WHERE   TPO.YEAR in (@ANO) 
                AND    TPO.MONTH<=8-- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                and oc.porIsIntegrated=3
        group by loc.Region

        UNION

        SELECT	@ANOM1 'Ano',
                loc.Region collate Modern_Spanish_CI_AS 'Region',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP_CAg'	    ,
                SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD_CAg'	    ,
                count(distinct porid) 'CantOC_CAg'
        FROM	[10.34.71.227].DM_Transaccional_2022.DBO.THOrdenesCompra AS OC     
                inner JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimTiempo AS TPO		ON OC.IDFechaEnvioOC=TPO.DateKey 
                inner join [10.34.71.227].DM_Transaccional_2022.dbo.DimComprador c			on c.IDUnidaddeCompra=oc.IDUnidaddeCompra 
                inner join [10.34.71.227].DM_Transaccional_2022.DBO.diminstitucion i		on i.entcode=c.entCode 
                inner JOIN [10.34.71.227].DM_Transaccional_2022.DBO.DimLocalidad as loc		ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
        WHERE   TPO.YEAR in (@ANOM1) 
                AND    TPO.MONTH<=8-- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                and oc.porIsIntegrated=3
        group by loc.Region              

        ORDER BY 'Region',
		'Ano' DESC
        '''
    return q


#------------------------#
# Montos OOPP Nacionales #
#------------------------#
def QueryOOPPNacional():
    q = '''
        SELECT TOP 3
        i.entCode,
        NombreInstitucion Institución,
        (sum(MontoUSD + ImpuestoUSD)) Monto_Bruto_USD,
        (sum(Montoclp + Impuestoclp)) Monto_Bruto_CLP

        FROM [DM_Transaccional].[dbo].[THOrdenesCompra] oc inner join
        [DM_Transaccional]..DimTiempo t on t.DateKey=oc.IDFechaEnvioOC left join
        [DM_Transaccional].dbo.DimComprador c on c.IDUnidaddeCompra=oc.IDUnidaddeCompra left join
        [DM_Transaccional].dbo.THOportunidadesNegocio opn on opn.rbhCode=oc.rbhCode left join
        [DM_Transaccional]..diminstitucion i on i.entcode=c.entCode 
        left JOIN DM_Transaccional..DimLocalidad as loc	  ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
        where
        t.Year in (2023)
        AND T.MONTH <=8
        group by
        i.entcode, NombreInstitucion
        order by sum(MontoUSD + ImpuestoUSD) desc, entCode
        '''
    return q


#----------------------------------------------#
# Montos totales transados por sector nacional #
#----------------------------------------------#
def QuerySectorNacional():
    q = '''
        SELECT 	Sector
                , SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD'
                , SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP'
                , count(oc.porid) CantOC
                
        FROM  DM_Transaccional..THOrdenesCompra AS OC     
            inner JOIN DM_Transaccional..DimTiempo AS TPO			ON OC.IDFechaEnvioOC=TPO.DateKey
            left JOIN DM_Transaccional..DimComprador AS C			ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
            left JOIN DM_Transaccional..DimInstitucion AS I			ON C.entCode = I.entCode   
            LEFT JOIN DM_Transaccional..DimSector AS S				ON S.IdSector = I.IdSector

        WHERE   TPO.YEAR in (2023) 
            AND	TPO.MONTH<=8

        GROUP BY  SECTOR

        order by Monto_Bruto_USD DESC
        '''
    return q


#-------------------#
# Regiones del país #
#-------------------#
def queryRegiones():
    q = '''
        SELECT DISTINCT [Region]
        FROM [DM_Transaccional].[dbo].[DimLocalidad]
        WHERE Region NOT IN ('Sin información', 'Extranjero')
        '''
    return q
