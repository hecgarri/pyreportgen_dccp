

		DECLARE @MESI INT, @MESF  INT
        SET @MESI = 1--''' +str(mi)+ '''
        SET @MESF = 11--''' +str(mf)+ '''

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
					WHEN 2 THEN 'MiPyme'
					WHEN 3 THEN 'MiPyme'
					WHEN 4 THEN 'MiPyme'
					WHEN 5 THEN 'MiPyme'
					ELSE CASE TMNa.tamanonombre
						WHEN 'Grande'THEN 'Grande'
						ELSE 'MiPyme'
					END
				END							'Tmn'
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
				,SEC.Sector 'Sec'
                
--------------------------------------------------------------------------------
        FROM  
            DM_Transaccional..THOrdenesCompra AS OC     
            INNER JOIN DM_Transaccional..DimTiempo AS TPO           ON OC.IDFechaEnvioOC = TPO.DateKey 
            LEFT JOIN DM_Transaccional..DimComprador AS COMP			ON OC.IDUnidaddeCompra = COMP.IDUnidaddeCompra 
            LEFT JOIN DM_Transaccional..DimLocalidad	 AS LOC		ON COMP.IDLocalidadUnidaddeCompra = LOC.IDLocalidad
				-------------------
				--HASTA AQUÍ BASE--
				-------------------

			LEFT JOIN DM_Transaccional..DimProveedor AS PROV			ON PROV.IDSucursal=OC.IDSucursal
			LEFT JOIN DM_Transaccional..THTamanoProveedor AS TMN1		ON TMN1.entcode = PROV.entCode AND AñoTributario=2021
			LEFT JOIN Estudios..TamanoProveedorNuevos20230802 AS TMNa	ON PROV.entcode = TMNa.entCode
            LEFT JOIN DM_Transaccional..DimInstitucion AS INS      ON COMP.entCode = INS.entCode
			LEFT JOIN DM_Transaccional.dbo.THOportunidadesNegocio AS LIC ON OC.rbhCode = LIC.rbhCode
            LEFT JOIN DM_Transaccional..DimSector AS SEC		ON SEC.IdSector = INS.IdSector
			
		

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
				,INS.NombreInstitucion
				,PROV.RazonSocialSucursal
				,PROV.RUTSucursal
				,OC.CodigoOC 
				,OC.Link 
				,ISNULL(LIC.NombreAdq, OC.NombreOC)
				,SEC.Sector
				
		ORDER BY Region, TPO.Year


		
        --    LEFT JOIN DM_Transaccional..DimProcedenciaOC AS PRO		ON PRO.IDProcedenciaOC = OC.IDProcedenciaOC