#################################
###   Querys para generador   ###
#################################
from datetime import datetime

#cabiar uso de parametros en query, del 'WHEN' a 'SET'

# retorna str que apunta a base de dato según año indicado
def bbddAno(ano):
    bbdd = ''
    if ano == datetime.now().year or ano == 2024: #actualizar
        bbdd = '[10.34.71.202].[DM_Transaccional].[dbo].'
    else: bbdd = '[DM_Transaccional_dev].[dbo].'
    return bbdd


#0000----------------------------------#
#0000 Query con datos a nivel regional #
#0000----------------------------------#
def theQueryReg(mi, mf): #Tiene Tamaño, Modalidad, Institución
    q = '''
		DECLARE @MESI INT, @MESF  INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''

        SELECT	TPO.YEAR							'Ano'
                ,loc.region							'Region'
                , SUM(OC.MontoUSD+OC.ImpuestoUSD)	'USD'
                , SUM(OC.MontoCLP+OC.ImpuestoCLP)	'CLP'
                , SUM(OC.MontoCLF+OC.ImpuestoCLF)	'CLF'
                , count(OC.porid)					'OC'
				-------------------
				--HASTA AQUÍ BASE--
				-------------------
				, CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					ELSE 'MiPyme'
				    END				as			'Tmn'
				, (CASE OC.porisintegrated
					WHEN 3 THEN 'Compra Ágil'
						else (case  OC.IDProcedenciaOC
								WHEN 703 THEN 'Convenio Marco'
								WHEN 701 THEN 'Licitación Pública'
								WHEN 1401 THEN 'Licitación Pública'
								WHEN 702 THEN 'Licitación Privada'
								ELSE 'Trato Directo'
							END)
					END) 'Mod'
				,INS.NombreInstitucion 'Ins'
				,PROV.RazonSocialSucursal 'Prv' --NombreSucursal
				,PROV.RUTSucursal	'PrvID'
				,OC.CodigoOC 'OCod'
				,OC.Link 'OLink'
				,ISNULL(LIC.NombreAdq, OC.NombreOC) 'Mtv'
				,CASE sec.SECTOR WHEN 'Gob. Central, Universidades' THEN 'Gob. Central, Ues.' 
		                else sec.SECTOR end  'Sec'
                
--------------------------------------------------------------------------------
        FROM  DM_Transaccional_dev..THOrdenesCompra AS OC     
            INNER JOIN DM_Transaccional_dev..DimTiempo AS TPO           ON OC.IDFechaEnvioOC = TPO.DateKey 
            LEFT JOIN DM_Transaccional_dev..DimComprador AS COMP			ON OC.IDUnidaddeCompra = COMP.IDUnidaddeCompra 
            LEFT JOIN DM_Transaccional_dev..DimLocalidad	 AS LOC		ON COMP.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
				-------------------
				--HASTA AQUÍ BASE--
				-------------------

			LEFT JOIN DM_Transaccional_dev..DimProveedor AS PROV			ON PROV.IDSucursal=OC.IDSucursal
			LEFT JOIN DM_Transaccional_dev..THTamanoProveedor AS TMN1		ON TMN1.entcode = PROV.entCode AND AñoTributario=(tpo.year -1) 
            LEFT JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion] AS INS      ON COMP.entCode = INS.entCode
			LEFT JOIN DM_Transaccional_dev.dbo.THOportunidadesNegocio AS LIC ON OC.rbhCode = LIC.rbhCode
            LEFT JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimSector] AS SEC		ON SEC.IdSector = INS.IdSector
			
		

--------------------------------------------------			
        WHERE   TPO.YEAR in (2023)
            AND	TPO.MONTH>= @MESI
            AND TPO.MONTH<= @MESF

        GROUP BY  Region
				,TPO.YEAR				
				-------------------
				--HASTA AQUÍ BASE--
				-------------------
				,CASE TMN1.idTamano
					WHEN 1 THEN 'Grande'
					ELSE 'MiPyme'
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
				,INS.NombreInstitucion
				,PROV.RazonSocialSucursal
				,PROV.RUTSucursal
				,OC.CodigoOC 
				,OC.Link 
				,ISNULL(LIC.NombreAdq, OC.NombreOC)
				,SEC.Sector
				
		ORDER BY Region, TPO.Year
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
                  --  , isnull(a.idtamano,'SinDato') end tamano0
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
                left join [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and a.AñoTributario=(tpo.year -1) 
                LEFT JOIN  DM_Transaccional..DimTamanoProveedor AS TP ON TP.IdTamano=a.IdTamano
                left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion] AS I      ON C.entCode = I.entCode
                left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimSector] AS S           ON S.IdSector = I.IdSector
                left JOIN DM_Transaccional..DimLocalidad as loc      ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

                --Dado que existe comparacion anual, debe tomarse con año 1 y año -1
                WHERE   TPO.YEAR in(2023
                                    --,2022
                                    )
                        AND TPO.MONTH >= @MESI AND TPO.MONTH <= @MESF
                GROUP BY TPO.YEAR,TPO.MONTH, 
               -- isnull(a.idtamano,'SinDato')  ,
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
            ,isnull(a.idtamano,'SinDato') end tamano0
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
            left join [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and a.AñoTributario=(tpo.year -1) 
            LEFT JOIN  DM_Transaccional..DimTamanoProveedor AS TP ON TP.IdTamano=a.IdTamano
            left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion] AS I      ON C.entCode = I.entCode
            left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimSector] AS S           ON S.IdSector = I.IdSector
            left JOIN DM_Transaccional..DimLocalidad as loc      ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
            
            --Dado que existe comparacion anual, debe tomarse con año 1 y año -1
            WHERE   TPO.YEAR in(2023,2022)
                    AND TPO.MONTH >= @MESI AND TPO.MONTH <= @MESF
            GROUP BY TPO.YEAR,TPO.MONTH, isnull(a.idtamano,'SinDato') ,
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
            ,case  tamano0 when '1' then 'Grande'
                else 'MiPyme'
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
            , case  tamano0 when '1' then 'Grande'
                else 'MiPyme'
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
            DM_Transaccional_dev..THOrdenesCompra AS OC     
            inner JOIN DM_Transaccional_dev..DimTiempo AS TPO           ON OC.IDFechaEnvioOC=TPO.DateKey 
            LEFT JOIN DM_Transaccional_dev..DimProcedenciaOC AS PRO		ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC
            left JOIN DM_Transaccional_dev..DimComprador AS C			ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra 
            left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion] AS I			ON C.entCode = I.entCode   
            LEFT JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimSector] AS S				ON S.IdSector = I.IdSector
            left JOIN DM_Transaccional_dev..DimLocalidad as loc			ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad

        WHERE   TPO.YEAR in (2022,2023)
            AND		TPO.MONTH>= @MESI
            AND     TPO.MONTH<= @MESF
        GROUP BY  Region,
                    TPO.YEAR

        order by Region
                ,TPO.YEAR
        '''
    return q


#0000--------------------------------------------------------------#
#0000 PROTO: Montos totales transados por sector por región (meses y año) #
#0000--------------------------------------------------------------#
def querySectorRegion(mi, mf, ano):
    #bbdd = bbddAno(ano)
    
    q = '''
		DECLARE @MESI INT, @MESF  INT, @ANO INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
		SET @ANO = ''' +str(ano)+ '''

        SELECT	TPO.YEAR							'Ano'
                ,LOC.region							'Region'
                , SUM(OC.MontoUSD+OC.ImpuestoUSD)	'USD'
                , SUM(OC.MontoCLP+OC.ImpuestoCLP)	'CLP'
                , SUM(OC.MontoCLF+OC.ImpuestoCLF)	'CLF'
                , COUNT(OC.porid)					'OC'
				,CASE sec.SECTOR WHEN 'Gob. Central, Universidades' THEN 'Gob. Central, Ues.' 
		            else sec.SECTOR end 	'Sec'

        FROM  
            [DM_Transaccional_dev].dbo.THOrdenesCompra			AS OC     
            INNER JOIN [DM_Transaccional_dev].dbo.DimTiempo		AS TPO	ON OC.IDFechaEnvioOC = TPO.DateKey 
            LEFT JOIN [DM_Transaccional_dev].dbo.DimComprador		AS COMP	ON OC.IDUnidaddeCompra = COMP.IDUnidaddeCompra 
            LEFT JOIN [DM_Transaccional_dev].dbo.DimLocalidad		AS LOC	ON COMP.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
            LEFT JOIN [10.34.71.202].[DM_Transaccional].dbo.DimInstitucion	AS INS	ON INS.entCode = COMP.entCode
            LEFT JOIN [10.34.71.202].[DM_Transaccional].dbo.DimSector		AS SEC	ON SEC.IdSector = INS.IdSector
		
        WHERE   TPO.YEAR in (@ANO)
            AND	TPO.MONTH>= @MESI
            AND TPO.MONTH<= @MESF
			AND LOC.region IS NOT NULL

        GROUP BY  Region
				,TPO.YEAR
				,SEC.Sector
				
		ORDER BY Region
				,TPO.YEAR
        '''
    return q


#oooo-----------------------------------------------#
#oooo Montos totales transados por rubro por región #
#oooo-----------------------------------------------#
def queryRubroRegion(mi, mf, top): # No cuuenta OC al dividir necesariamente por item comprado
    q = ''' 
        DECLARE @MESI INT, @MESF  INT, @TOP INT
        SET @MESI = ''' +str(mi)+ '''
        SET @MESF = ''' +str(mf)+ '''
        SET @TOP = ''' +str(top)+ '''
        SELECT * 
        FROM (SELECT
				TPO.Year			'Ano'
				,LOC.Region			'Region'
				,rubro.RubroN1		'Rub'
				,SUM(OCL.MontoUSD)  'USD'
				,SUM(OCL.MontoCLP)	'CLP'
				,SUM(OCL.MontoCLF)	'CLF'
				,0					'OC'	-- columna para estandarizar tabla a requerimientos del generador de reportes
				,Rank() OVER (PARTITION BY LOC.Region ORDER BY  SUM(OCL.MontoUSD) DESC ) 
								AS	Rank
			FROM  
				DM_Transaccional_dev..THOrdenesCompra				  AS OC     
				LEFT JOIN DM_Transaccional_dev..DimTiempo			  AS TPO	ON OC.IDFechaEnvioOC=TPO.DateKey 
				LEFT JOIN DM_Transaccional_dev..DimComprador		  AS C      ON OC.IDUnidaddeCompra = C.IDUnidaddeCompra
				LEFT JOIN DM_Transaccional_dev..DimLocalidad		  AS LOC	ON C.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
				LEFT JOIN  DM_Transaccional_dev..THOrdenesCompraLinea AS OCL	ON OCL.porID=OC.porID
				LEFT JOIN DM_Transaccional_dev..DimProducto			  AS PRODUC ON PRODUC.IDProducto=OCL.IDProducto 
				LEFT JOIN DM_Transaccional_dev..DimRubro			  AS RUBRO	ON RUBRO.IdRubro=PRODUC.IdRubro
			WHERE   TPO.Year = 2023
				AND  TPO.Month BETWEEN @MESI AND @MESF
			GROUP BY LOC.Region, TPO.Year, RUBRO.RubroN1
					) AS RS
		WHERE Rank <= @TOP
		ORDER BY Region ASC, Rank ASC 
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

        SELECT	TPO.Year [Ano],
                loc.Region 'Region',
                SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP_CAg'
                ,SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD_CAg'
                ,SUM(OC.MontoCLF+OC.ImpuestoCLF) 'MONTOCLF_CAg'
                ,COUNT(DISTINCT porid) 'CantOC_CAg'
        FROM	dm_transaccional_dev..THOrdenesCompra AS OC     
                inner JOIN dm_transaccional_dev..DimTiempo AS TPO		ON OC.IDFechaEnvioOC=TPO.DateKey 
                inner join [dm_transaccional_dev].dbo.DimComprador c	on C.IDUnidaddeCompra=oc.IDUnidaddeCompra
                inner join [10.34.71.202].[dm_transaccional].[dbo].[diminstitucion] i		on i.entcode=c.entCode 
                inner JOIN dm_transaccional_dev..DimLocalidad as loc	ON C.IDLocalidadUnidaddeCompra =  LOC.IDLocalidad
        WHERE   TPO.YEAR in (@ANO,@ANOM1) 
                AND    TPO.MONTH >= @MESI
                AND    TPO.MONTH <= @MESF-- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
                and oc.porIsIntegrated=3
        group by TPO.Year,loc.Region

        ORDER BY 'Region',
		TPO.Year DESC
        '''
    return q

#!
#-------------------------#
# Query Top OC por Región #
#-------------------------#
def queryOrdenCompraRegionTop(top, mi, mf):
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
                        end 'Tamano',
                        [10.34.71.202].DM_Transaccional.dbo.DimInstitucion.NombreInstitucion,
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

                FROM	[10.34.71.202].DM_Transaccional.dbo.DimInstitucion INNER JOIN
                        [10.34.71.202].DM_Transaccional.dbo.DimComprador ON DM_Transaccional.dbo.DimInstitucion.entCode = DM_Transaccional.dbo.DimComprador.entCode INNER JOIN
                        DM_Transaccional.dbo.DimLocalidad ON DM_Transaccional.dbo.DimComprador.IDLocalidadUnidaddeCompra = DM_Transaccional.dbo.DimLocalidad.IDLocalidad INNER JOIN
                        DM_Transaccional.dbo.THOrdenesCompra ON DM_Transaccional.dbo.DimComprador.IDUnidaddeCompra = DM_Transaccional.dbo.THOrdenesCompra.IDUnidaddeCompra INNER JOIN
                        DM_Transaccional.dbo.DimProveedor ON DM_Transaccional.dbo.DimProveedor.IDSucursal = DM_Transaccional.dbo.THOrdenesCompra.IDSucursal INNER JOIN
                        --DM_Transaccional.dbo.DimTamanoProveedor ON DM_Transaccional.dbo.DimTamanoProveedor.IdTamano = DM_Transaccional.dbo.DimProveedor.IdTamano INNER JOIN
                        DM_Transaccional.dbo.DimTiempo ON DM_Transaccional.dbo.DimTiempo.DateKey =  DM_Transaccional.dbo.THOrdenesCompra.IDFechaEnvioOC LEFT JOIN
                        DM_Transaccional.dbo.THOportunidadesNegocio ON DM_Transaccional.dbo.THOportunidadesNegocio.rbhCode = DM_Transaccional.dbo.THOrdenesCompra.rbhCode
                        left join   [DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=[DM_Transaccional].[dbo].[dimproveedor].entCode and AñoTributario=(DM_Transaccional.dbo.DimTiempo.year-1)
                        --left join Estudios.dbo.TamanoProveedorNuevos20230809 b on [DM_Transaccional].[dbo].[dimproveedor].entcode=b.entCode
                
                where	DM_Transaccional.dbo.dimtiempo.Year = 2023
                        and DM_Transaccional.dbo.dimtiempo.month between @MESI and @MESF
                ) rs

            WHERE Rank <= ''' +str(top)+ '''

            ORDER BY Region ASC, Rank ASC
        '''
    return q


#O------------------------------------------#
#O Query OC por Región (q grande como base) #
#O------------------------------------------#
def queryOrdenCompraRegion(mi, mf):
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
				,INS.NombreInstitucion				'Ins'
				,PROV.RazonSocialSucursal			'Prv' --puede ser .NombreSucursal
				,PROV.RUTSucursal					'PrvID'
				,OC.CodigoOC						'OCod'
				,OC.Link							'OLink'
				,ISNULL(LIC.NombreAdq, OC.NombreOC) 'Mtv'
                
        FROM  
            DM_Transaccional..THOrdenesCompra	AS OC     
            INNER JOIN DM_Transaccional..DimTiempo				  AS TPO  ON OC.IDFechaEnvioOC = TPO.DateKey 
            LEFT JOIN DM_Transaccional..DimComprador			  AS COMP ON OC.IDUnidaddeCompra = COMP.IDUnidaddeCompra 
            LEFT JOIN DM_Transaccional..DimLocalidad			  AS LOC  ON COMP.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
			LEFT JOIN DM_Transaccional..DimProveedor			  AS PROV ON OC.IDSucursal = PROV.IDSucursal
            LEFT JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion]			  AS INS  ON COMP.entCode = INS.entCode
			LEFT JOIN DM_Transaccional.dbo.THOportunidadesNegocio AS LIC  ON OC.rbhCode = LIC.rbhCode
	
        WHERE   TPO.YEAR in (2023)
			AND TPO.MONTH>= @MESI
			AND TPO.MONTH<= @MESF

        GROUP BY  Region
				,TPO.YEAR				
				,INS.NombreInstitucion
				,PROV.RazonSocialSucursal
				,PROV.RUTSucursal
				,OC.CodigoOC 
				,OC.Link 
				,ISNULL(LIC.NombreAdq, OC.NombreOC)
				
		ORDER BY Region, TPO.Year
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
        WHERE   TPO.YEAR in (@ANO,@ANOM1) -- and (oc.OCEXCEPCIONAL=0 or oc.OCEXCEPCIONAL is null ) 
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
        [10.34.71.202].[DM_Transaccional].[dbo].[diminstitucion] i on i.entcode=c.entCode 
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
            left JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimInstitucion] AS I			ON C.entCode = I.entCode   
            LEFT JOIN [10.34.71.202].[DM_Transaccional].[dbo].[DimSector] AS S				ON S.IdSector = I.IdSector

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
                case when a.idTamano is not null then 
                        case a.idTamano when 1 then 'Grande'
                            else 'Mipyme'
                        end 
                    end as [Tamano]
                ,SUM(OC.MontoUSD+OC.ImpuestoUSD) 'MONTOUSD'
                , SUM(OC.MontoCLF+OC.ImpuestoCLF) 'MONTOCLF'
                ,SUM(OC.MontoCLP+OC.ImpuestoCLP) 'MONTOCLP'
                ,count(distinct oc.porid) CantOC
                ,COUNT(DISTINCT P.RUTSucursal) 'CantProveedores'
                    --  into #aux1

        FROM	DM_Transaccional_dev..THOrdenesCompra AS OC     
                inner JOIN DM_Transaccional_dev..DimTiempo AS TPO                     ON OC.IDFechaEnvioOC=TPO.DateKey 
                inner join [DM_Transaccional_dev].[dbo].[dimproveedor] p on p.orgCode=oc.IDSucursal
                left join   [10.34.71.202].[DM_Transaccional].[dbo].[THTamanoProveedor] a on a.entcode=p.entCode and AñoTributario=2022

        WHERE   TPO.YEAR in (2023) and TPO.MONTH >= @MESI AND TPO.MONTH <= @MESO

        group by case 
                    when a.idTamano is not null then 
                    case a.idTamano when 1 then 'Grande'
                        else 'Mipyme'
                    end
                end
        '''
    return q


#0000-------------------#
#0000 Regiones del país #
#0000-------------------#
def queryRegiones():
    q = '''
        SELECT DISTINCT [Region]
        FROM [DM_Transaccional_dev].[dbo].[DimLocalidad]
        WHERE Region NOT IN ('Sin información', 'Extranjero')
        '''
    return q


#0000-------------------#
#0000 Sectores del país #
#0000-------------------#
def querySectores():
    q = '''
        SELECT CASE SECTOR WHEN 'Gob. Central, Universidades' THEN 'Gob. Central, Ues.' 
		            else SECTOR end 	'Sec'
        FROM [10.34.71.202].[DM_Transaccional].[dbo].DimSector
        WHERE Sector <> 'SINDATO'
        ORDER BY Sector 
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
 