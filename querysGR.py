#########################
# Querys para generador #
#########################




def QueryReg ():
    q = ''';with temp as (
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

def queryTotalRegional():
    q = '''SELECT		TPO.YEAR 'Año'
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
                    ,TPO.YEAR'''
    return q