#################################
###   Querys para generador   ###
#################################


#cabiar uso de parametros en query, del 'WHEN' a 'SET'

#----------------------------------#
# Query con datos a nivel regional #
#----------------------------------#
def theQueryReg(mi, mf): #Tiene Tamaño y Modalidad
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
        SELECT	TPO.YEAR							'Ano'
                ,loc.region							'Region'
                , SUM(OC.MontoUSD+OC.ImpuestoUSD)	'USD'
                , SUM(OC.MontoCLP+OC.ImpuestoCLP)	'CLP'
                , SUM(OC.MontoCLF+OC.ImpuestoCLF)	'CLF'
                , count(oc.porid)					'OC'
				-------------------
				--HASTA AQUÍ BASE--
				-------------------
				, CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					WHEN 2 THEN 'MiPyme'
					WHEN 3 THEN 'MiPyme'
					WHEN 4 THEN 'MiPyme'
					WHEN 5 THEN 'MiPyme'
					ELSE CASE TMNa.tamanonombre
						WHEN 'Grande'THEN 'Grande'
						ELSE 'MiPyme'
					END
				END							'Tamano'
				, (CASE OC.porisintegrated
					WHEN 3 THEN 'Compra Ágil'
						else (case  OC.IDProcedenciaOC
								WHEN 703 THEN 'Convenio Marco'
								WHEN 701 THEN 'Licitación Pública'
								WHEN 1401 THEN 'Licitación Pública'
								WHEN 702 THEN 'Licitación Privada'
								ELSE 'Trato Directo'
							END)
					END) 'Modalidad'
                
        FROM  
            DM_Transaccional..THOrdenesCompra AS OC     
            INNER JOIN DM_Transaccional..DimTiempo AS TPO           ON OC.IDFechaEnvioOC = TPO.DateKey 
            LEFT JOIN DM_Transaccional..DimComprador AS COMP			ON OC.IDUnidaddeCompra = COMP.IDUnidaddeCompra 
            LEFT JOIN DM_Transaccional..DimLocalidad	 AS LOC		ON COMP.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
				-------------------
				--HASTA AQUÍ BASE--
				-------------------

			LEFT JOIN DM_Transaccional..DimProveedor AS PROV			ON PROV.IDSucursal=OC.IDSucursal
			LEFT JOIN DM_Transaccional.dbo.THTamanoProveedor AS TMN1		ON TMN1.entcode = PROV.entCode AND AñoTributario=2021
			LEFT JOIN Estudios.dbo.TamanoProveedorNuevos20230802 AS TMNa	ON PROV.entcode = TMNa.entCode

			
            /*LEFT JOIN DM_Transaccional..DimInstitucion AS I		ON C.entCode = I.entCode   
            LEFT JOIN DM_Transaccional..DimSector AS S				ON S.IdSector = I.IdSector

			LEFT JOIN DM_Transaccional..DimTamanoProveedor AS TP	ON TP.IdTamano=a.IdTamano*/
			
        WHERE   TPO.YEAR in (2023)
            AND		TPO.MONTH>= @MESI
            AND     TPO.MONTH<= @MESF
        GROUP BY  Region
				,TPO.YEAR				
				-------------------
				--HASTA AQUÍ BASE--
				-------------------
				,CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					WHEN 2 THEN 'MiPyme'
					WHEN 3 THEN 'MiPyme'
					WHEN 4 THEN 'MiPyme'
					WHEN 5 THEN 'MiPyme'
					ELSE CASE TMNa.tamanonombre
						WHEN 'Grande'THEN 'Grande'
						ELSE 'MiPyme'
					END
				END			
				, (CASE OC.porisintegrated
					WHEN 3 THEN 'Compra Ágil'
						else (case  OC.IDProcedenciaOC
								WHEN 703 THEN 'Convenio Marco'
								WHEN 701 THEN 'Licitación Pública'
								WHEN 1401 THEN 'Licitación Pública'
								WHEN 702 THEN 'Licitación Privada'
								ELSE 'Trato Directo'
							END)
					END)

		ORDER BY Region
        '''
    return q


#-----------------------------------------------------------#
# Queri monto instituciones regional (QueryReg() como base) #
#-----------------------------------------------------------#
def queryInstitucionRegion(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
        ;with temp as (
                SELECT TPO.YEAR 'Ano'
                    , tpo.Month as 'Mes'
                    , LOC.Region
                    , (CASE oc1.porisintegrated WHEN 3 THEN 'Compra Ágil'
                            else (
                                case  OC1.IDProcedenciaOC
                                    WHEN 703 THEN 'Convenio Marco'
                                    WHEN 701 THEN 'Licitación Pública'
                                    WHEN 1401 THEN 'Licitación Pública'
                                    WHEN 702 THEN 'Licitación Privada'
                                    ELSE 'Trato Directo' 
                                END)
                        END) 'ProcedenciaOC'
                    , I.NombreInstitucion
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
                WHERE   TPO.YEAR in(2023
                                    --,2022
                                    )
                        AND TPO.MONTH >= @MESI AND TPO.MONTH <= @MESF
                GROUP BY TPO.YEAR,TPO.MONTH, 
                case when a.idTamano is not null then cast(a.idTamano as nvarchar)
                            else isnull(b.tamanonombre,'SinDato') end ,
                            I.NombreInstitucion,  LOC.Region,
                    (CASE OC1.porisintegrated WHEN 3 THEN 'Compra Ágil'
                            else (case  OC1.IDProcedenciaOC
                                            WHEN 703 THEN 'Convenio Marco'
                                            WHEN 701 THEN 'Licitación Pública'
                                            WHEN 1401 THEN 'Licitación Pública'
                                            WHEN 702 THEN 'Licitación Privada'
                                            ELSE 'Trato Directo' END)END)
                )
        Select 	Region
                , NombreInstitucion
                ,sum(MONTOUSD) MONTOUSD,
                sum(MONTOCLF) MONTOCLF,
                sum(MONTOCLP) MONTOCLP,
                sum(CantOC) CantOC,
                sum(CantProveedores)      'CantProveedores'
        from	temp
        group by	Region
                    , NombreInstitucion
        ORDER BY	Region
                    , sum(MONTOUSD) DESC
        '''
    return q


#--------------------------------------------------------#
# Query con datos a nivel nacional. Solo no tiene Region #
#--------------------------------------------------------#
def QueryTotal(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
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
                    AND TPO.MONTH >= @MESI AND TPO.MONTH <= @MESF
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
def queryTotalRegion(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
        SELECT		TPO.YEAR 'Año'
                    ,loc.region 'Region'
                    , SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD'
                    , SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP'
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
            AND		TPO.MONTH>= @MESI
            AND     TPO.MONTH<= @MESF
        GROUP BY  Region,
                    TPO.YEAR
            --SECTOR

        UNION

        SELECT		TPO.YEAR 'Año'
                    ,loc.region collate Modern_Spanish_CI_AS 'Region'
                    ,SUM(OC.MontoUSD+OC.ImpuestoUSD) 'Monto_Bruto_USD'
                    ,SUM(OC.MontoCLP+OC.ImpuestoCLP) 'Monto_Bruto_CLP'
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
            AND		TPO.MONTH>= @MESI
            AND     TPO.MONTH<= @MESF
        GROUP BY  Region
                ,TPO.YEAR

        order by Region
                ,TPO.YEAR
        '''
    return q


#------------------------------------------------#
# Montos totales transados por sector por región #
#------------------------------------------------#
def querySectorRegion(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

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
            AND TPO.MONTH >= @MESI
            AND	TPO.MONTH <= @MESF
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
def queryRubroRegion(mi, mf):
    q = ''' 
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
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
                AND  TPO.month between @MESI and @MESF
            GROUP BY LOC.Region,rubro.RubroN1
                    )rs WHERE Rank <= 3 ORDER BY Region ASC, Rank ASC 
        '''
    return q


#-----------------------------------------#
# Query CA por región periodo 2023 y 2022 #
#-----------------------------------------#
def queryCompraAgilRegion(mi, mf):
    q = '''
        DECLARE @ANO Int,  @ANOM1 Int, @MESI INT, @MESF  INT
        SET @ANO = 2023
        SET @ANOM1 = @ANO - 1
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

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
                AND    TPO.MONTH >= @MESI
                AND    TPO.MONTH <= @MESF-- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
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
                AND    TPO.MONTH >= @MESI
                AND    TPO.MONTH <= @MESF-- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                and oc.porIsIntegrated=3
        group by loc.Region              

        ORDER BY 'Region',
		'Ano' DESC
        '''
    return q

#!
#-------------------------#
# Query Top OC por Región #
#-------------------------#
def queryOrdenCompraRegion(top, mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
        SELECT *
            FROM (
                SELECT 
                        case 
                            when a.idTamano is not null then 
                                case a.idTamano
                                    when 1 then 'Grande'
                                    else 'Mipyme'
                                end
                            else 
                                case isnull(b.Tamano,5) 
                                    when 1 then 'Grande'
                                    else 'Mipyme' 
                                end 
                        end 'Tamano',
                        DM_Transaccional.dbo.DimInstitucion.NombreInstitucion,
                        DM_Transaccional.dbo.DimComprador.NombreUnidaddeCompra,
                        DM_Transaccional.dbo.DimLocalidad.Region,
                        DM_Transaccional.dbo.DimProveedor.NombreSucursal,
                        --DM_Transaccional.dbo.DimTamanoProveedor.Tamano,
                        DM_Transaccional.dbo.THOrdenesCompra.CodigoOC,
                        DM_Transaccional.dbo.THOrdenesCompra.NombreOC,
                        DM_Transaccional.dbo.THOportunidadesNegocio.NombreAdq as NombreLic,
                        ISNULL(DM_Transaccional.dbo.THOportunidadesNegocio.NombreAdq, DM_Transaccional.dbo.THOrdenesCompra.NombreOC) as MotivoCompra,
                        round(DM_Transaccional.dbo.THOrdenesCompra.MontoUSD + DM_Transaccional.dbo.THOrdenesCompra.ImpuestoUSD,0) USD_BRUTO,
                        DM_Transaccional.dbo.THOrdenesCompra.MontoCLP + DM_Transaccional.dbo.THOrdenesCompra.ImpuestoCLP CLP_BRUTO,
                        Rank()
                        OVER (Partition BY DM_Transaccional.dbo.DimLocalidad.Region
                            ORDER BY DM_Transaccional.dbo.THOrdenesCompra.MontoUSD DESC ) AS Rank

                FROM	DM_Transaccional.dbo.DimInstitucion INNER JOIN
                        DM_Transaccional.dbo.DimComprador ON DM_Transaccional.dbo.DimInstitucion.entCode = DM_Transaccional.dbo.DimComprador.entCode INNER JOIN
                        DM_Transaccional.dbo.DimLocalidad ON DM_Transaccional.dbo.DimComprador.IDLocalidadUnidaddeCompra = DM_Transaccional.dbo.DimLocalidad.IDLocalidad INNER JOIN
                        DM_Transaccional.dbo.THOrdenesCompra ON DM_Transaccional.dbo.DimComprador.IDUnidaddeCompra = DM_Transaccional.dbo.THOrdenesCompra.IDUnidaddeCompra INNER JOIN
                        DM_Transaccional.dbo.DimProveedor ON DM_Transaccional.dbo.DimProveedor.IDSucursal = DM_Transaccional.dbo.THOrdenesCompra.IDSucursal INNER JOIN
                        --DM_Transaccional.dbo.DimTamanoProveedor ON DM_Transaccional.dbo.DimTamanoProveedor.IdTamano = DM_Transaccional.dbo.DimProveedor.IdTamano INNER JOIN
                        DM_Transaccional.dbo.DimTiempo ON DM_Transaccional.dbo.DimTiempo.DateKey =  DM_Transaccional.dbo.THOrdenesCompra.IDFechaEnvioOC LEFT JOIN
                        DM_Transaccional.dbo.THOportunidadesNegocio ON DM_Transaccional.dbo.THOportunidadesNegocio.rbhCode = DM_Transaccional.dbo.THOrdenesCompra.rbhCode
                        left join   [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=[DM_Transaccional].[dbo].[dimproveedor].entCode and AñoTributario=2021
                        left join Estudios.dbo.TamanoProveedorNuevos20230809 b on [DM_Transaccional].[dbo].[dimproveedor].entcode=b.entCode
                
                where	DM_Transaccional.dbo.dimtiempo.Year = 2023
                        and DM_Transaccional.dbo.dimtiempo.month between @MESI and @MESF
                ) rs

            WHERE Rank <= ''' +str(top)+ '''

            ORDER BY Region ASC, Rank ASC
        '''
    return q


#---------------------------------------------#
# Query Proveedores Regiones(reg compradores) #
#---------------------------------------------#
def queryProveedoresRegiones(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT, @ANO Int, @ANOM1 Int
        SET @ANO = 2023
        SET @ANOM1 = @ANO - 1
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

        SELECT @ANO 'Ano', 
				Loc.Region,
				Prov.RazonSocialSucursal,
				Prov.RUTSucursal,
                SUM(OC.MontoCLF+OC.ImpuestoCLF) 'CLF', 
                SUM(OC.MontoUSD+OC.ImpuestoUSD) 'USD',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'CLP', 
                COUNT(OC.MontoCLF) 'CANTIDADOC'	

        FROM	DM_Transaccional..THOrdenesCompra AS OC INNER JOIN 
                DM_Transaccional..DimTiempo AS TPO	ON OC.IDFechaEnvioOC=TPO.DateKey LEFT JOIN
				DM_Transaccional.dbo.DimProveedor AS Prov ON OC.IDSucursal = Prov.IDSucursal LEFT JOIN
				DM_Transaccional.dbo.DimComprador AS Comp ON OC.IDUnidaddeCompra = Comp.IDUnidaddeCompra LEFT JOIN
				DM_Transaccional.dbo.DimLocalidad AS Loc ON Loc.IDLocalidad = Comp.IDLocalidadUnidaddeCompra
				WHERE   TPO.YEAR in (@ANO) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                AND tpo.Month >= @MESI
                AND tpo.Month <= @MESF
		
		GROUP BY Region, Prov.RUTSucursal, Prov.RazonSocialSucursal
        ORDER BY ANO DESC, Loc.Region, SUM(OC.MontoUSD+OC.ImpuestoUSD) DESC, Prov.RUTSucursal
        '''
    return q


#--------------------------#
# Query Totales Nacionales #
#--------------------------#
def QueryTotalesNacionales(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT, @ANO Int, @ANOM1 Int
        SET @ANO = 2023
        SET @ANOM1 = @ANO - 1
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

        SELECT @ANO 'ANO', 
                SUM(OC.MontoCLF+OC.ImpuestoCLF) 'MONTOCLF', 
                SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP', 
                COUNT(OC.MontoCLF) 'CANTIDADOC'	
        FROM DM_Transaccional..THOrdenesCompra AS OC LEFT JOIN 
                DM_Transaccional..DimTiempo AS TPO	ON OC.IDFechaEnvioOC=TPO.DateKey 
        WHERE   TPO.YEAR in (@ANO) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                AND tpo.Month >= @MESI
                AND tpo.Month <= @MESF

        UNION

        SELECT @ANOM1 'ANO', 
                SUM(OC.MontoCLF+OC.ImpuestoCLF) 'MONTOCLF', 
                SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP', 
                COUNT(OC.MontoCLF) 'CANTIDADOC'	
        FROM [10.34.71.227].DM_Transaccional_2022.dbo.THOrdenesCompra AS OC LEFT JOIN 
                [10.34.71.227].DM_Transaccional_2022.dbo.DimTiempo AS TPO    ON OC.IDFechaEnvioOC=TPO.DateKey 
        WHERE   TPO.YEAR in (@ANOM1) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                AND tpo.Month >= @MESI
                AND tpo.Month <= @MESF

        ORDER BY ANO DESC
        '''
    return q 


#------------------------#
# Montos OOPP Nacionales #
#------------------------#
def QueryOrgnismosPublicoNacional(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

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
        AND T.MONTH >= @MESI
        AND T.MONTH <= @MESF
        group by
        i.entcode, NombreInstitucion
        order by sum(MontoUSD + ImpuestoUSD) desc, entCode
        '''
    return q


#----------------------------------------------#
# Montos totales transados por sector nacional #
#----------------------------------------------#
def QuerySectorNacional(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

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
            AND	TPO.MONTH >= @MESI
            AND	TPO.MONTH <= @MESF

        GROUP BY  SECTOR

        order by Monto_Bruto_USD DESC
        '''
    return q


#----------------------------------------------#
# Totales Proveedores Grande y Mipyme Nacional #
#----------------------------------------------#
def QueryTotalProveedoresNacional(mi, mf):
    q = '''
        DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

        SELECT       
                case 
                    when a.idTamano is not null then 
                        case a.idTamano when 1 then 'Grande'
                            else 'Mipyme'
                        end
                    else 
                        case isnull(b.Tamano,5) 
                            when 1 then 'Grande'
                            else 'Mipyme'
                        end
                end 'Tamano'
                ,SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD'
                , SUM(OC.MontoCLF+OC.ImpuestoCLF) 'MONTOCLF'
                ,SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP'
                ,count(distinct oc.porid) CantOC
                ,COUNT(DISTINCT P.RUTSucursal) 'CantProveedores'
                    --  into #aux1

        FROM	DM_Transaccional..THOrdenesCompra AS OC     
                inner JOIN DM_Transaccional..DimTiempo AS TPO                     ON OC.IDFechaEnvioOC=TPO.DateKey 
                inner join [DM_Transaccional].[dbo].[dimproveedor] p on p.orgCode=oc.IDSucursal
                left join   [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and AñoTributario=2021
                left join Estudios.dbo.TamanoProveedorNuevos20230809 b on p.entcode=b.entCode

        WHERE   TPO.YEAR in (2023) and TPO.MONTH >= @MESI AND TPO.MONTH <= @MESO

        group by case 
                    when a.idTamano is not null then 
                    case a.idTamano when 1 then 'Grande'
                        else 'Mipyme'
                    end
                else 
                    case isnull(b.Tamano,5) 
                        when 1 then 'Grande'
                        else 'Mipyme'
                    end 
                end
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


#------------------------------------#
# Query Cantidad Proveedores Aquiles #
#------------------------------------#

    # ###########################################################################
    # ### Query para obtener número total de proveedores que participan en MP ###
    # ###########################################################################
    # QProv = pd.read_sql(con = conn_AQ, params= FechaProv2, sql = '''
    
    #    /*Proveedores involucrados en una OC*/
    #    SELECT DISTINCT C.entCode AS Transan
    #    FROM         
    #    DCCPProcurement..prcPOHeader A with(nolock)INNER JOIN
    #    DCCPPlatform..gblOrganization B with(nolock) ON A.porSellerOrganization = B.orgCode INNER JOIN
    #    DCCPPlatform..gblEnterprise C with(nolock) ON B.orgEnterprise = C.entCode
    #    WHERE     (A.porBuyerStatus IN (4, 5, 6, 7, 12)) 
    #    AND year(A.porSendDate) =  ? 
    #    and ( month(A.porSendDate)>= ? AND month(A.porSendDate)<= ? )
       
    #    UNION 
       
    #    /*Proveedores que han participado emitiendo una oferta*/
       
    #    SELECT DISTINCT C.orgEnterprise as Transan
    #    FROM         
    #    DCCPProcurement..prcBIDQuote A with(nolock)INNER JOIN
    #    DCCPProcurement..prcRFBHeader B with(nolock) ON A.bidRFBCode = B.rbhCode INNER JOIN
    #    DCCPPlatform..gblOrganization C with(nolock) ON A.bidOrganization = C.orgCode
    #    WHERE     (A.bidDocumentStatus IN (3, 4, 5)) 
    #    AND year(A.bidEconomicIssueDate) = ? 
    #    and (month(A.bidEconomicIssueDate)>= ? AND month(A.bidEconomicIssueDate)<= ? )
    
    #    UNION
       
    #    /*Proveedores a los que se les ha solicitado una cotización*/
       
    #    SELECT  DISTINCT (B.orgEnterprise) AS Transan     
    #      FROM [DCCPProcurement].[dbo].[prcPOCotizacion] A
    #      INNER JOIN DCCPPlatform..gblOrganization B ON
    #      A.proveedorRut=B.orgTaxID INNER JOIN
    #      DCCPProcurement..prcPOHeader C ON
    #      A.porId = C.porID
    #      WHERE year(C.porSendDate) = ? 
    #    and (month(C.porSendDate)>= ? AND month(C.porSendDate)<= ?)
       
    #    UNION
    #    /*MÁS LOS RUT NO REGISTRADOS EN MERCADO PÚBLICO Y QUE SON DEL PERÍODO DE REFERENCIA*/
       
    #    SELECT  
    #    DISTINCT A.proveedorRut   
    #      FROM [DCCPProcurement].[dbo].[prcPOCotizacion] A INNER JOIN
    #       DCCPProcurement..prcPOHeader C ON
    #      A.porId = C.porID
    #      WHERE A.proveedorRut NOT IN (
    #    SELECT  
    #    DISTINCT A.proveedorRut     
    #      FROM [DCCPProcurement].[dbo].[prcPOCotizacion] A
    #      INNER  JOIN DCCPPlatform..gblOrganization B ON
    #      A.proveedorRut=B.orgTaxID) AND 
    #      year(C.porSendDate) = ? and 
    #      (month(C.porSendDate)>= ? AND month(C.porSendDate)<= ?)''')

    ####  ESTA ES EL OUTPUT DE LA QUERY ANTERIOR PERO YA EJECUTADA EN SQL SERVER, PORQUE LA ANTERIOR DEMORA MUCHO ####
    #
    # QProv=pd.read_sql(con=conn_AQ, sql='''
    #     SELECT *
    #     FROM estudios.dbo.ProveedoresTransando20230811
    #     ''')

    # QProvTrans = pd.read_sql(con = conn_AQ,  sql = '''
    # SELECT DISTINCT C.entCode AS Transan
    # FROM         
    # DCCPProcurement..prcPOHeader A with(nolock)INNER JOIN
    # DCCPPlatform..gblOrganization B with(nolock) ON A.porSellerOrganization = B.orgCode INNER JOIN
    # DCCPPlatform..gblEnterprise C with(nolock) ON B.orgEnterprise = C.entCode
    # WHERE     (A.porBuyerStatus IN (4, 5, 6, 7, 12)) 
    # AND year(A.porSendDate) =  2023 
    # and (month(A.porSendDate)>=1 AND month(A.porSendDate)<= 8 )
    # ''')
 