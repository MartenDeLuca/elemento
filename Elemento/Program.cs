using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Timers;
using System.Net.Mail;
using System.Net.Mime;
using static System.Net.Mime.MediaTypeNames;

namespace Elemento
{
    internal class Program
    {
        private static String conexion = "";
        private static String empresa_a = "";
        private static String empresa_b = "";
        private static String emailDestinatarios = "";
        private static String talonario_pedido = "";
		private static String se_borro_ultima_vez = "";

		/*
		//Configuro para que se ejecute automaticamente
		[DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        private static System.Timers.Timer timer_informacion;
        private static System.Timers.Timer timer_informacion_error;
		*/
        static void Main(string[] args)
        {
			ejecutoProcesoMain();
			/*
			var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
            Timer_informacion();
            Console.ReadLine();
            //Console.ReadKey();            
			*/
        }

		/*
        private static void Timer_informacion()
        {
            // Crea un temporizador con un intervalo 1 minutos
            timer_informacion = new System.Timers.Timer(60000);
            // Conecta el evento transcurrido para el temporizador.
            timer_informacion.Elapsed += proceso_informacion;
            timer_informacion.AutoReset = true;
            timer_informacion.Enabled = true;
        }

        private static void Timer_informacion_error()
        {
            // Crea un temporizador con un intervalo de un minuto
            timer_informacion_error = new System.Timers.Timer(300000);
            // Conecta el evento transcurrido para el temporizador.
            timer_informacion_error.Elapsed += revisar_error_estructura;
            timer_informacion_error.AutoReset = true;
            timer_informacion_error.Enabled = true;
        }

        private static void revisar_error_estructura(Object source, ElapsedEventArgs e)
        {
            timer_informacion_error.Stop();
            timer_informacion_error.Dispose();
            Timer_informacion();
        }
		
        private static void proceso_informacion(Object source, ElapsedEventArgs e)
        {
            ejecutoProcesoMain();
            Console.ReadKey();
        }
		*/

		private static void ejecutoProcesoMain()
        {
            List<String> hayErrores = new List<String>();
            try
            {
                conexion = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
                empresa_a = ConfigurationManager.ConnectionStrings["empresa_a"].ConnectionString;
                empresa_b = ConfigurationManager.ConnectionStrings["empresa_b"].ConnectionString;
                talonario_pedido = ConfigurationManager.ConnectionStrings["talonario_pedido"].ConnectionString;
                emailDestinatarios = ConfigurationManager.ConnectionStrings["emailDestinatarios"].ConnectionString;
                using (var con = new SqlConnection(conexion))
                {

                    con.Open();

                    SqlTransaction transaction;
                    transaction = con.BeginTransaction("SampleTransaction");
                    SqlCommand sqlComando = new SqlCommand();
                    sqlComando.Connection = con;
                    sqlComando.Transaction = transaction;
                    String talonario = "", n_comp = "", t_comp = "", insertar = "", ncomp_in_s = "", tcomp_in_s = "", tiene_errores = "", cod_client = "", cod_direccion_entrega = "", id_direccion_entrega = "";
                    DataTable tablaConsulta = new DataTable(), tablaConsulta2 = new DataTable();

					try
                    {
						String sql = @"SELECT distinct sta14.talonario, sta14.n_comp, sta14.t_comp, sta14.tcomp_in_s, sta14.ncomp_in_s, case when STA14_ERRORES.N_COMP IS NULL then 'No' else 'Si' end tiene_errores
							FROM [" + empresa_b + @"]..gva106
							INNER JOIN (SELECT DISTINCT sta14.TALONARIO, sta14.n_comp, sta14.t_comp, sta14.tcomp_in_s, sta14.ncomp_in_s FROM [" + empresa_b + @"]..sta14) sta14 on sta14.NCOMP_IN_S = GVA106.NCOMP_IN_S and sta14.TCOMP_IN_S = GVA106.TCOMP_IN_S 
							LEFT JOIN (SELECT DISTINCT sta14.TALONARIO, sta14.n_comp, sta14.t_comp, sta14.tcomp_in_s, sta14.ncomp_in_s FROM [" + empresa_a + @"]..sta14) sta14_a on sta14.N_COMP = sta14_a.N_COMP and sta14.T_COMP = sta14_a.T_COMP
							LEFT JOIN [" + empresa_a + @"]..STA14_ERRORES on STA14_ERRORES.N_COMP = STA14.N_COMP collate Latin1_General_BIN AND STA14_ERRORES.T_COMP = STA14.T_COMP collate Latin1_General_BIN
							WHERE sta14.T_COMP = 'REM' and sta14_a.TALONARIO IS NULL AND GVA106.TALON_PED = '" + talonario_pedido + "'";
						sqlComando.CommandText = sql;
                        sqlComando.CommandType = CommandType.Text;
                        sqlComando.ExecuteNonQuery();
                        SqlDataAdapter SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
                        tablaConsulta = new DataTable();
                        SqlAdaptadorDatos.Fill(tablaConsulta);
						if (tablaConsulta.Rows.Count > 0)
						{
							foreach (DataRow fila in tablaConsulta.Rows)
							{
								talonario = fila["talonario"].ToString();
								n_comp = fila["n_comp"].ToString();
								t_comp = fila["t_comp"].ToString();
								tcomp_in_s = fila["tcomp_in_s"].ToString();
								ncomp_in_s = fila["ncomp_in_s"].ToString();
								tiene_errores = fila["tiene_errores"].ToString();

								hayErrores = new List<String>();

								sqlComando.CommandText =
								@"SELECT distinct STA20.COD_ARTICU
								from [" + empresa_b + @"]..STA20
								INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA20.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA20.TCOMP_IN_S
								LEFT JOIN [" + empresa_a + @"]..sta11 ON sta11.COD_ARTICU = STA20.COD_ARTICU
								WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA11.COD_ARTICU IS NULL";
								sqlComando.CommandType = CommandType.Text;
								sqlComando.ExecuteNonQuery();
								SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
								tablaConsulta2 = new DataTable();
								SqlAdaptadorDatos.Fill(tablaConsulta2);
								if (tablaConsulta2.Rows.Count > 0)
								{
									String articulos_errores = "";
									foreach (DataRow fila2 in tablaConsulta2.Rows)
									{
										articulos_errores += fila2["COD_ARTICU"].ToString()+", ";
									}
									if(articulos_errores != "")
                                    {
										articulos_errores = articulos_errores.Substring(0, articulos_errores.Length - 2);
                                    }

									hayErrores.Add("Los artículos "+articulos_errores+" del remito NO existen en la empresa A");
								}

								sqlComando.CommandText =
								@"SELECT distinct STA20.COD_DEPOSI 
								from [" + empresa_b + @"]..STA20
								INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA20.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA20.TCOMP_IN_S
								LEFT JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA20.COD_DEPOSI
								WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA22.COD_SUCURS IS NULL";
								sqlComando.CommandType = CommandType.Text;
								sqlComando.ExecuteNonQuery();
								SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
								tablaConsulta2 = new DataTable();
								SqlAdaptadorDatos.Fill(tablaConsulta2);
								if (tablaConsulta2.Rows.Count > 0)
								{
									String deposito_error = "";
									foreach (DataRow fila2 in tablaConsulta2.Rows)
									{
										deposito_error += fila2["COD_DEPOSI"].ToString();
									}
									hayErrores.Add("El deposito "+deposito_error+" del remito NO existen en la empresa A");
								}

								sqlComando.CommandText =
								@"SELECT distinct GVA14_b.RAZON_SOCI
								from [" + empresa_b + @"]..STA14
								INNER JOIN [" + empresa_b + @"]..GVA14 GVA14_b ON GVA14_b.COD_CLIENT = STA14.COD_PRO_CL
								LEFT JOIN [" + empresa_a + @"]..GVA14 GVA14_a ON GVA14_b.RAZON_SOCI = GVA14_a.RAZON_SOCI
								WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND GVA14_a.cod_client IS NULL";
								sqlComando.CommandType = CommandType.Text;
								sqlComando.ExecuteNonQuery();
								SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
								tablaConsulta2 = new DataTable();
								SqlAdaptadorDatos.Fill(tablaConsulta2);
								if (tablaConsulta2.Rows.Count > 0)
								{
									String cod_client_error = "";
									foreach (DataRow fila2 in tablaConsulta2.Rows)
									{
										cod_client_error += fila2["RAZON_SOCI"].ToString();
									}
									hayErrores.Add("El cliente "+ cod_client_error + " del remito NO existen en la empresa A");
                                }
                                else
                                {
									sqlComando.CommandText =
									@"SELECT distinct DIRECCION_ENTREGA.COD_DIRECCION_ENTREGA, GVA14_a.COD_CLIENT
									from [" + empresa_b + @"]..STA14
									LEFT JOIN [" + empresa_b + @"]..DIRECCION_ENTREGA ON DIRECCION_ENTREGA.ID_DIRECCION_ENTREGA = STA14.ID_DIRECCION_ENTREGA
									INNER JOIN [" + empresa_b + @"]..GVA14 GVA14_b ON GVA14_b.COD_CLIENT = STA14.COD_PRO_CL
									LEFT JOIN [" + empresa_a + @"]..GVA14 GVA14_a ON GVA14_b.RAZON_SOCI = GVA14_a.RAZON_SOCI								
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"'";
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();
									SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
									tablaConsulta2 = new DataTable();
									SqlAdaptadorDatos.Fill(tablaConsulta2);
									if (tablaConsulta2.Rows.Count > 0)
                                    {
										foreach (DataRow fila2 in tablaConsulta2.Rows)
										{
											cod_client = fila2["COD_CLIENT"].ToString();
											cod_direccion_entrega = fila2["COD_DIRECCION_ENTREGA"].ToString();

											sqlComando.CommandText =
											@"SELECT distinct DIRECCION_ENTREGA.ID_DIRECCION_ENTREGA
											FROM [" + empresa_a + @"]..DIRECCION_ENTREGA
											WHERE DIRECCION_ENTREGA.COD_CLIENTE = '" + cod_client + @"' AND DIRECCION_ENTREGA.COD_DIRECCION_ENTREGA = '" + cod_direccion_entrega + @"'";
											sqlComando.CommandType = CommandType.Text;
											sqlComando.ExecuteNonQuery();
											SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
											DataTable tablaConsulta3 = new DataTable();
											SqlAdaptadorDatos.Fill(tablaConsulta3);
											if (tablaConsulta3.Rows.Count > 0)
											{
												foreach (DataRow fila3 in tablaConsulta3.Rows)
												{
													id_direccion_entrega = fila3["ID_DIRECCION_ENTREGA"].ToString();
												}
                                            }
                                            else
                                            {
												hayErrores.Add("La dirección con codigo "+cod_direccion_entrega+" del cliente " + cod_client + " del remito NO existen en la empresa A");
											}
										}

									}

								}

								if (hayErrores.Count == 0)
								{
									insertar =
									@"INSERT INTO [" + empresa_a + @"].[dbo].[STA14]
									([FILLER]
									,[COD_PRO_CL]
									,[COTIZ]
									,[ESTADO_MOV]
									,[EXPORTADO]
									,[EXP_STOCK]
									,[FECHA_ANU]
									,[FECHA_MOV]
									,[HORA]
									,[LISTA_REM]
									,[LOTE]
									,[LOTE_ANU]
									,[MON_CTE]
									,[MOTIVO_REM]
									,[N_COMP]
									,[N_REMITO]
									,[NCOMP_IN_S]
									,[NCOMP_ORIG]
									,[NRO_SUCURS]
									,[OBSERVACIO]
									,[SUC_ORIG]
									,[T_COMP]
									,[TALONARIO]
									,[TCOMP_IN_S]
									,[TCOMP_ORIG]
									,[USUARIO]
									,[COD_TRANSP]
									,[HORA_COMP]
									,[ID_A_RENTA]
									,[DOC_ELECTR]
									,[COD_CLASIF]
									,[AUDIT_IMP]
									,[IMP_IVA]
									,[IMP_OTIMP]
									,[IMPORTE_BO]
									,[IMPORTE_TO]
									,[DIFERENCIA]
									,[SUC_DESTIN]
									,[T_DOC_DTE]
									,[LEYENDA1]
									,[LEYENDA2]
									,[LEYENDA3]
									,[LEYENDA4]
									,[LEYENDA5]
									,[DCTO_CLIEN]
									,[T_INT_ORI]
									,[N_INT_ORI]
									,[FECHA_INGRESO]
									,[HORA_INGRESO]
									,[USUARIO_INGRESO]
									,[TERMINAL_INGRESO]
									,[IMPORTE_TOTAL_CON_IMPUESTOS]
									,[CANTIDAD_KILOS]
									,[ID_DIRECCION_ENTREGA]
									,[IMPORTE_GRAVADO]
									,[IMPORTE_EXENTO]
									,[ID_STA13]
									,[NRO_SUCURSAL_DESTINO_REMITO])
									select 
									[STA14].[FILLER]
									,'"+cod_client+@"' [COD_PRO_CL]
									,[STA14].[COTIZ]
									,[STA14].[ESTADO_MOV]
									,[STA14].[EXPORTADO]
									,[STA14].[EXP_STOCK]
									,[STA14].[FECHA_ANU]
									,[STA14].[FECHA_MOV]
									,[STA14].[HORA]
									,[STA14].[LISTA_REM]
									,[STA14].[LOTE]
									,[STA14].[LOTE_ANU]
									,[STA14].[MON_CTE]
									,[STA14].[MOTIVO_REM]
									,[STA14].[N_COMP]
									,[STA14].[N_REMITO]
									,RIGHT('00000000' + convert(varchar(1000),((SELECT ISNULL(MAX(NCOMP_IN_S),0) FROM [" + empresa_a + @"]..[STA14] sta14max WHERE sta14max.TCOMP_IN_S = sta14.tcomp_in_s) + (ROW_NUMBER() OVER (ORDER BY STA14.NCOMP_IN_S)))), 8) [NCOMP_IN_S]
									,[STA14].[NCOMP_ORIG]
									,[STA14].[NRO_SUCURS]
									,[STA14].[OBSERVACIO]
									,[STA14].[SUC_ORIG]
									,[STA14].[T_COMP]
									,[STA14].[TALONARIO]
									,[STA14].[TCOMP_IN_S]
									,[STA14].[TCOMP_ORIG]
									,[STA14].[USUARIO]
									,[STA14].[COD_TRANSP]
									,[STA14].[HORA_COMP]
									,[STA14].[ID_A_RENTA]
									,[STA14].[DOC_ELECTR]
									,[STA14].[COD_CLASIF]
									,[STA14].[AUDIT_IMP]
									,[STA14].[IMP_IVA]
									,[STA14].[IMP_OTIMP]
									,[STA14].[IMPORTE_BO]
									,[STA14].[IMPORTE_TO]
									,[STA14].[DIFERENCIA]
									,[STA14].[SUC_DESTIN]
									,[STA14].[T_DOC_DTE]
									,[STA14].[LEYENDA1]
									,[STA14].[LEYENDA2]
									,[STA14].[LEYENDA3]
									,[STA14].[LEYENDA4]
									,[STA14].[LEYENDA5]
									,[STA14].[DCTO_CLIEN]
									,[STA14].[T_INT_ORI]
									,[STA14].[N_INT_ORI]
									,[STA14].[FECHA_INGRESO]
									,[STA14].[HORA_INGRESO]
									,[STA14].[USUARIO_INGRESO]
									,[STA14].[TERMINAL_INGRESO]
									,[STA14].[IMPORTE_TOTAL_CON_IMPUESTOS]
									,[STA14].[CANTIDAD_KILOS]
									,'"+ id_direccion_entrega + @"' [ID_DIRECCION_ENTREGA]
									,[STA14].[IMPORTE_GRAVADO]
									,[STA14].[IMPORTE_EXENTO]
									,[STA14].[ID_STA13]
									,[STA14].[NRO_SUCURSAL_DESTINO_REMITO]
									from [" + empresa_b + @"]..sta14
									LEFT JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.TALONARIO = sta14.TALONARIO  
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND sta14_destino.id_sta14 is null";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();


									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA14TY]
									([FILLER]
									,[TCOMP_IN_S]
									,[NCOMP_IN_S]
									,[IMG_TYP]
									,[TALONARIO])
									SELECT
										STA14TY.[FILLER]
										,sta14_destino.[TCOMP_IN_S]
										,sta14_destino.[NCOMP_IN_S]
										,STA14TY.[IMG_TYP]
										,STA14TY.[TALONARIO]
									FROM [" + empresa_b + @"].[dbo].[STA14TY]
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA14TY.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA14TY.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									LEFT JOIN [" + empresa_a + @"]..STA14TY STA14TY_destino ON STA14_destino.NCOMP_IN_S = STA14TY_destino.NCOMP_IN_S AND STA14_destino.TCOMP_IN_S = STA14TY_destino.TCOMP_IN_S
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA14TY_destino.NCOMP_IN_S IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA20]
									([FILLER]
									,[CAN_EQUI_V]
									,[CANT_DEV]
									,[CANT_OC]
									,[CANT_PEND]
									,[CANT_SCRAP]
									,[CANTIDAD]
									,[COD_ARTICU]
									,[COD_DEPOSI]
									,[DEPOSI_DDE]
									,[EQUIVALENC]
									,[FECHA_MOV]
									,[N_ORDEN_CO]
									,[N_RENGL_OC]
									,[N_RENGL_S]
									,[NCOMP_IN_S]
									,[PLISTA_REM]
									,[PPP_EX]
									,[PPP_LO]
									,[PRECIO]
									,[PRECIO_REM]
									,[TCOMP_IN_S]
									,[TIPO_MOV]
									,[COD_CLASIF]
									,[CANT_FACTU]
									,[DCTO_FACTU]
									,[CANT_DEV_2]
									,[CANT_PEND_2]
									,[CANTIDAD_2]
									,[CANT_FACTU_2]
									,[ID_MEDIDA_STOCK_2]
									,[ID_MEDIDA_STOCK]
									,[ID_MEDIDA_VENTAS]
									,[ID_MEDIDA_COMPRA]
									,[UNIDAD_MEDIDA_SELECCIONADA]
									,[PRECIO_REMITO_VENTAS]
									,[CANT_OC_2]
									,[RENGL_PADR]
									,[COD_ARTICU_KIT]
									,[PROMOCION]
									,[PRECIO_ADICIONAL_KIT]
									,[TALONARIO_OC]
									,[ID_STA11]
									,[ID_STA14]
									,[COD_DEPOSI_INGRESO])

									SELECT 
									[STA20].[FILLER]
									,[STA20].[CAN_EQUI_V]
									,[STA20].[CANT_DEV]
									,[STA20].[CANT_OC]
									,[STA20].[CANT_PEND]
									,[STA20].[CANT_SCRAP]
									,[STA20].[CANTIDAD]
									,[STA20].[COD_ARTICU]
									,[STA20].[COD_DEPOSI]
									,[STA20].[DEPOSI_DDE]
									,[STA20].[EQUIVALENC]
									,[STA20].[FECHA_MOV]
									,[STA20].[N_ORDEN_CO]
									,[STA20].[N_RENGL_OC]
									,[STA20].[N_RENGL_S]
									,[sta14_destino].[NCOMP_IN_S]
									,[STA20].[PLISTA_REM]
									,[STA20].[PPP_EX]
									,[STA20].[PPP_LO]
									,[STA20].[PRECIO]
									,[STA20].[PRECIO_REM]
									,[STA20].[TCOMP_IN_S]
									,[STA20].[TIPO_MOV]
									,[STA20].[COD_CLASIF]
									,[STA20].[CANT_FACTU]
									,[STA20].[DCTO_FACTU]
									,[STA20].[CANT_DEV_2]
									,[STA20].[CANT_PEND_2]
									,[STA20].[CANTIDAD_2]
									,[STA20].[CANT_FACTU_2]
									,[STA20].[ID_MEDIDA_STOCK_2]
									,[STA20].[ID_MEDIDA_STOCK]
									,[STA20].[ID_MEDIDA_VENTAS]
									,[STA20].[ID_MEDIDA_COMPRA]
									,[STA20].[UNIDAD_MEDIDA_SELECCIONADA]
									,[STA20].[PRECIO_REMITO_VENTAS]
									,[STA20].[CANT_OC_2]
									,[STA20].[RENGL_PADR]
									,[STA20].[COD_ARTICU_KIT]
									,[STA20].[PROMOCION]
									,[STA20].[PRECIO_ADICIONAL_KIT]
									,[STA20].[TALONARIO_OC]
									,sta11.[ID_STA11]
									,sta14_destino.[ID_STA14]
									,[STA20].[COD_DEPOSI_INGRESO]
									from [" + empresa_b + @"]..STA20
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA20.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA20.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA20.COD_DEPOSI
									INNER JOIN [" + empresa_a + @"]..sta11 ON sta11.COD_ARTICU = STA20.COD_ARTICU
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									LEFT JOIN [" + empresa_a + @"]..STA20 STA20_destino ON STA14_destino.NCOMP_IN_S = STA20_destino.NCOMP_IN_S AND STA14_destino.TCOMP_IN_S = STA20_destino.TCOMP_IN_S
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA20_destino.ID_STA20 IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA08]
									([FILLER]
									,[ADUANA]
									,[COD_PROVEE]
									,[COMENTARIO]
									,[FECHA]
									,[FECHA_VTO]
									,[ID_CARPETA]
									,[N_COMP]
									,[N_DESPACHO]
									,[N_PARTIDA]
									,[PAIS]
									,[T_COMP]
									,[PARTIDA_DESC_ADICIONAL_1]
									,[PARTIDA_DESC_ADICIONAL_2]
									,[ID_CPA01]
									,[ID_STA14]
									,[ID_STA13])
									SELECT DISTINCT
									[STA08].[FILLER]
									,[STA08].[ADUANA]
									,[STA08].[COD_PROVEE]
									,[STA08].[COMENTARIO]
									,[STA08].[FECHA]
									,[STA08].[FECHA_VTO]
									,[STA08].[ID_CARPETA]
									,[STA08].[N_COMP]
									,[STA08].[N_DESPACHO]
									,[STA08].[N_PARTIDA]
									,[STA08].[PAIS]
									,[STA08].[T_COMP]
									,[STA08].[PARTIDA_DESC_ADICIONAL_1]
									,[STA08].[PARTIDA_DESC_ADICIONAL_2]
									,null ID_CPA01
									,[sta14_destino].[ID_STA14]
									,NULL [ID_STA13]
									FROM [" + empresa_b + @"]..STA08
									INNER JOIN [" + empresa_b + @"]..STA09 ON STA09.N_PARTIDA = STA08.N_PARTIDA		
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA09.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA09.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA09.COD_DEPOSI
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									LEFT JOIN [" + empresa_a + @"]..STA08 STA08_DESTINO ON STA08_DESTINO.N_PARTIDA = STA08.N_PARTIDA
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA08_DESTINO.N_PARTIDA is null";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA09]
									([FILLER]
									,[CANTIDAD]
									,[COD_ARTICU]
									,[COD_DEPOSI]
									,[N_PARTIDA]
									,[N_RENGL_S]
									,[NCOMP_IN_S]
									,[TCOMP_IN_S]
									,[CANTIDAD_2]
									,[CANT_DEV]
									,[CANT_DEV_2])
									SELECT
									[STA09].[FILLER]
									,[STA09].[CANTIDAD]
									,[STA09].[COD_ARTICU]
									,[STA09].[COD_DEPOSI]
									,[STA09].[N_PARTIDA]
									,[STA09].[N_RENGL_S]
									,[sta14_destino].[NCOMP_IN_S]
									,[STA09].[TCOMP_IN_S]
									,[STA09].[CANTIDAD_2]
									,[STA09].[CANT_DEV]
									,[STA09].[CANT_DEV_2]
									FROM [" + empresa_b + @"]..STA09
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA09.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA09.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA09.COD_DEPOSI
									INNER JOIN [" + empresa_a + @"]..sta11 ON sta11.COD_ARTICU = STA09.COD_ARTICU
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									LEFT JOIN [" + empresa_a + @"]..STA09 STA09_destino ON STA14_destino.NCOMP_IN_S = STA09_destino.NCOMP_IN_S AND STA14_destino.TCOMP_IN_S = STA09_destino.TCOMP_IN_S
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA09_destino.ID_STA09 IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA06]
									([FILLER]
									,[COD_ARTICU]
									,[DESC1]
									,[DESC2]
									,[N_SERIE]
									,[COD_DEPOSI]
									,[COMENTARIO]
									,[N_PARTIDA])
									SELECT distinct
									[STA06].[FILLER]
									,[STA06].[COD_ARTICU]
									,[STA06].[DESC1]
									,[STA06].[DESC2]
									,[STA06].[N_SERIE]
									,[STA06].[COD_DEPOSI]
									,[STA06].[COMENTARIO]
									,[STA06].[N_PARTIDA]
									FROM [" + empresa_b + @"]..STA06
									INNER JOIN [" + empresa_b + @"]..STA07 ON STA07.N_PARTIDA = STA06.N_PARTIDA AND STA07.N_SERIE = STA06.N_SERIE AND STA07.COD_ARTICU = STA06.COD_ARTICU
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA07.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA07.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA07.COD_DEPOSI
									INNER JOIN [" + empresa_a + @"]..sta11 ON sta11.COD_ARTICU = STA07.COD_ARTICU
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									INNER JOIN [" + empresa_b + @"]..STA06 STA06_DESTINO ON STA06_DESTINO.N_PARTIDA = STA06.N_PARTIDA AND STA06_DESTINO.N_SERIE = STA06.N_SERIE AND STA06_DESTINO.COD_ARTICU = STA06.COD_ARTICU
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' AND STA06_DESTINO.N_SERIE IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();


									insertar = @"
									INSERT INTO [" + empresa_a + @"].[dbo].[STA07]
									([FILLER]
									,[COD_ARTICU]
									,[DESC1]
									,[DESC2]
									,[N_RENGL_S]
									,[N_SERIE]
									,[NCOMP_IN_S]
									,[TCOMP_IN_S]
									,[COD_DEPOSI]
									,[COMENTARIO]
									,[N_PARTIDA])
									SELECT distinct
									[STA07].[FILLER]
									,[STA07].[COD_ARTICU]
									,[STA07].[DESC1]
									,[STA07].[DESC2]
									,[STA07].[N_RENGL_S]
									,[STA07].[N_SERIE]
									,[sta14_destino].[NCOMP_IN_S]
									,[STA07].[TCOMP_IN_S]
									,[STA07].[COD_DEPOSI]
									,[STA07].[COMENTARIO]
									,[STA07].[N_PARTIDA]
									FROM [" + empresa_b + @"]..STA07
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = STA07.NCOMP_IN_S AND STA14.TCOMP_IN_S = STA07.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..STA22 ON STA22.COD_SUCURS = STA07.COD_DEPOSI
									INNER JOIN [" + empresa_a + @"]..sta11 ON sta11.COD_ARTICU = STA07.COD_ARTICU
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario
									LEFT JOIN [" + empresa_a + @"]..STA07 STA07_destino ON STA14_destino.NCOMP_IN_S = STA07_destino.NCOMP_IN_S AND STA14_destino.TCOMP_IN_S = STA07_destino.TCOMP_IN_S AND STA07_destino.N_SERIE = STA07.N_SERIE AND STA07_destino.N_RENGL_S = STA07.N_RENGL_S
									WHERE STA14.TALONARIO = '" + talonario + @"' AND STA14.N_COMP = '" + n_comp + @"' AND STA14.T_COMP = '" + t_comp + @"' and STA07_destino.NCOMP_IN_S IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();


									insertar = @"
									INSERT INTO  [" + empresa_a + @"].[dbo].[GVA21]
									   ([FILLER]
									   ,[APRUEBA]
									   ,[CIRCUITO]
									   ,[COD_CLIENT]
									   ,[COD_SUCURS]
									   ,[COD_TRANSP]
									   ,[COD_VENDED]
									   ,[COMENTARIO]
									   ,[COMP_STK]
									   ,[COND_VTA]
									   ,[COTIZ]
									   ,[ESTADO]
									   ,[EXPORTADO]
									   ,[FECHA_APRU]
									   ,[FECHA_ENTR]
									   ,[FECHA_PEDI]
									   ,[HORA_APRUE]
									   ,[ID_EXTERNO]
									   ,[LEYENDA_1]
									   ,[LEYENDA_2]
									   ,[LEYENDA_3]
									   ,[LEYENDA_4]
									   ,[LEYENDA_5]
									   ,[MON_CTE]
									   ,[N_LISTA]
									   ,[N_REMITO]
									   ,[NRO_O_COMP]
									   ,[NRO_PEDIDO]
									   ,[NRO_SUCURS]
									   ,[ORIGEN]
									   ,[PORC_DESC]
									   ,[REVISO_FAC]
									   ,[REVISO_PRE]
									   ,[REVISO_STK]
									   ,[TALONARIO]
									   ,[TALON_PED]
									   ,[TOTAL_PEDI]
									   ,[TIPO_ASIEN]
									   ,[MOTIVO]
									   ,[HORA]
									   ,[COD_CLASIF]
									   ,[ID_ASIENTO_MODELO_GV]
									   ,[TAL_PE_ORI]
									   ,[NRO_PE_ORI]
									   ,[FECHA_INGRESO]
									   ,[HORA_INGRESO]
									   ,[USUARIO_INGRESO]
									   ,[TERMINAL_INGRESO]
									   ,[FECHA_ULTIMA_MODIFICACION]
									   ,[HORA_ULTIMA_MODIFICACION]
									   ,[USUA_ULTIMA_MODIFICACION]
									   ,[TERM_ULTIMA_MODIFICACION]
									   ,[ID_DIRECCION_ENTREGA]
									   ,[ES_PEDIDO_WEB]
									   ,[WEB_ORDER_ID]
									   ,[FECHA_O_COMP]
									   ,[ACTIVIDAD_COMPROBANTE_AFIP]
									   ,[ID_ACTIVIDAD_EMPRESA_AFIP]
									   ,[TIPO_DOCUMENTO_PAGADOR]
									   ,[NUMERO_DOCUMENTO_PAGADOR]
									   ,[USUARIO_TIENDA]
									   ,[TIENDA]
									   ,[ORDER_ID_TIENDA]
									   ,[NRO_OC_COMP]
									   ,[TIENDA_QUE_VENDE]
									   ,[TOTAL_DESC_TIENDA]
									   ,[PORCEN_DESC_TIENDA]
									   /*,[ID_GVA01]
									   ,[ID_GVA10]
									   ,[ID_GVA14]
									   ,[ID_GVA23]
									   ,[ID_GVA24]
									   ,[ID_GVA38]
									   ,[ID_GVA43_TALON_PED]
									   ,[ID_GVA81]
									   ,[ID_SUCURSAL]
									   ,[USUARIO_TIENDA_VENDEDOR]
									   ,[ID_NEXO_PEDIDOS_ORDEN]
									   ,[METODO_EXPORTACION]
									   ,[NRO_SUCURSAL_DESTINO_PEDIDO]
									   ,[ID_MODELO_PEDIDO]*/)
									SELECT 
										GVA21.[FILLER]
									   ,GVA21.[APRUEBA]
									   ,GVA21.[CIRCUITO]
									   ,'"+cod_client+@"' [COD_CLIENT]
									   ,GVA21.[COD_SUCURS]
									   ,GVA21.[COD_TRANSP]
									   ,GVA21.[COD_VENDED]
									   ,GVA21.[COMENTARIO]
									   ,GVA21.[COMP_STK]
									   ,GVA21.[COND_VTA]
									   ,GVA21.[COTIZ]
									   ,GVA21.[ESTADO]
									   ,GVA21.[EXPORTADO]
									   ,GVA21.[FECHA_APRU]
									   ,GVA21.[FECHA_ENTR]
									   ,GVA21.[FECHA_PEDI]
									   ,GVA21.[HORA_APRUE]
									   ,GVA21.[ID_EXTERNO]
									   ,GVA21.[LEYENDA_1]
									   ,GVA21.[LEYENDA_2]
									   ,GVA21.[LEYENDA_3]
									   ,GVA21.[LEYENDA_4]
									   ,GVA21.[LEYENDA_5]
									   ,GVA21.[MON_CTE]
									   ,GVA21.[N_LISTA]
									   ,GVA21.[N_REMITO]
									   ,GVA21.[NRO_O_COMP]
									   ,GVA21.[NRO_PEDIDO]
									   ,GVA21.[NRO_SUCURS]
									   ,GVA21.[ORIGEN]
									   ,GVA21.[PORC_DESC]
									   ,GVA21.[REVISO_FAC]
									   ,GVA21.[REVISO_PRE]
									   ,GVA21.[REVISO_STK]
									   ,GVA21.[TALONARIO]
									   ,GVA21.[TALON_PED]
									   ,GVA21.[TOTAL_PEDI]
									   ,GVA21.[TIPO_ASIEN]
									   ,GVA21.[MOTIVO]
									   ,GVA21.[HORA]
									   ,GVA21.[COD_CLASIF]
									   ,GVA21.[ID_ASIENTO_MODELO_GV]
									   ,GVA21.[TAL_PE_ORI]
									   ,GVA21.[NRO_PE_ORI]
									   ,GVA21.[FECHA_INGRESO]
									   ,GVA21.[HORA_INGRESO]
									   ,GVA21.[USUARIO_INGRESO]
									   ,GVA21.[TERMINAL_INGRESO]
									   ,GVA21.[FECHA_ULTIMA_MODIFICACION]
									   ,GVA21.[HORA_ULTIMA_MODIFICACION]
									   ,GVA21.[USUA_ULTIMA_MODIFICACION]
									   ,GVA21.[TERM_ULTIMA_MODIFICACION]
									   ,GVA21.[ID_DIRECCION_ENTREGA]
									   ,GVA21.[ES_PEDIDO_WEB]
									   ,GVA21.[WEB_ORDER_ID]
									   ,GVA21.[FECHA_O_COMP]
									   ,GVA21.[ACTIVIDAD_COMPROBANTE_AFIP]
									   ,GVA21.[ID_ACTIVIDAD_EMPRESA_AFIP]
									   ,GVA21.[TIPO_DOCUMENTO_PAGADOR]
									   ,GVA21.[NUMERO_DOCUMENTO_PAGADOR]
									   ,GVA21.[USUARIO_TIENDA]
									   ,GVA21.[TIENDA]
									   ,GVA21.[ORDER_ID_TIENDA]
									   ,GVA21.[NRO_OC_COMP]
									   ,GVA21.[TIENDA_QUE_VENDE]
									   ,GVA21.[TOTAL_DESC_TIENDA]
									   ,GVA21.[PORCEN_DESC_TIENDA]
									   /*,GVA21.[ID_GVA01]
									   ,GVA21.[ID_GVA10]
									   ,GVA21.[ID_GVA14]
									   ,GVA21.[ID_GVA23]
									   ,GVA21.[ID_GVA24]
									   ,GVA21.[ID_GVA38]
									   ,GVA21.[ID_GVA43_TALON_PED]
									   ,GVA21.[ID_GVA81]
									   ,GVA21.[ID_SUCURSAL]
									   ,GVA21.[USUARIO_TIENDA_VENDEDOR]
									   ,GVA21.[ID_NEXO_PEDIDOS_ORDEN]
									   ,GVA21.[METODO_EXPORTACION]
									   ,GVA21.[NRO_SUCURSAL_DESTINO_PEDIDO]
									   ,GVA21.[ID_MODELO_PEDIDO]*/
									FROM [" + empresa_b + @"].[dbo].[GVA21]
									LEFT JOIN [" + empresa_b + @"]..GVA55 ON GVA55.TALON_PED = GVA21.TALON_PED AND GVA55.NRO_PEDIDO = GVA21.NRO_PEDIDO
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = gva55.T_COMP AND sta14_destino.N_COMP = gva55.N_COMP
									LEFT JOIN [" + empresa_a + @"]..GVA21 GVA21_destino ON GVA21_destino.NRO_PEDIDO = GVA21.NRO_PEDIDO AND GVA21_destino.TALON_PED = GVA21.TALON_PED										
									WHERE GVA55.N_COMP = '" + n_comp + @"' AND GVA55.T_COMP = '" + t_comp + @"' AND GVA21_destino.NRO_PEDIDO IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar =
									@"INSERT INTO [" + empresa_a + @"].[dbo].[GVA03]
									([FILLER]
									,[CAN_EQUI_V]
									,[CANT_A_DES]
									,[CANT_A_FAC]
									,[CANT_PEDID]
									,[CANT_PEN_D]
									,[CANT_PEN_F]
									,[COD_ARTICU]
									,[DESCUENTO]
									,[N_RENGLON]
									,[NRO_PEDIDO]
									,[PEN_REM_FC]
									,[PEN_FAC_RE]
									,[PRECIO]
									,[TALON_PED]
									,[COD_CLASIF]
									,[CANT_A_DES_2]
									,[CANT_A_FAC_2]
									,[CANT_PEDID_2]
									,[CANT_PEN_D_2]
									,[CANT_PEN_F_2]
									,[PEN_REM_FC_2]
									,[PEN_FAC_RE_2]
									,[ID_MEDIDA_VENTAS]
									,[ID_MEDIDA_STOCK_2]
									,[ID_MEDIDA_STOCK]
									,[UNIDAD_MEDIDA_SELECCIONADA]
									,[COD_ARTICU_KIT]
									,[RENGL_PADR]
									,[PROMOCION]
									,[PRECIO_ADICIONAL_KIT]
									,[KIT_COMPLETO]
									,[INSUMO_KIT_SEPARADO]
									,[PRECIO_LISTA]
									,[PRECIO_BONIF]
									,[DESCUENTO_PARAM]
									,[PRECIO_FECHA]
									,[FECHA_MODIFICACION_PRECIO]
									,[USUARIO_MODIFICACION_PRECIO]
									,[TERMINAL_MODIFICACION_PRECIO]
									,[ID_NEXO_PEDIDOS_RENGLON_ORDEN]
									,[CANT_A_DES_EXPORTADA]
									,[CANT_A_FAC_EXPORTADA]
									,[CANT_A_DES_2_EXPORTADA]
									,[CANT_A_FAC_2_EXPORTADA]
									,[COD_DEPOSI])
									SELECT
										GVA03.[FILLER]
										,GVA03.[CAN_EQUI_V]
										,GVA03.[CANT_A_DES]
										,GVA03.[CANT_A_FAC]
										,GVA03.[CANT_PEDID]
										,GVA03.[CANT_PEN_D]
										,GVA03.[CANT_PEN_F]
										,GVA03.[COD_ARTICU]
										,GVA03.[DESCUENTO]
										,GVA03.[N_RENGLON]
										,GVA03.[NRO_PEDIDO]
										,GVA03.[PEN_REM_FC]
										,GVA03.[PEN_FAC_RE]
										,GVA03.[PRECIO]
										,GVA03.[TALON_PED]
										,GVA03.[COD_CLASIF]
										,GVA03.[CANT_A_DES_2]
										,GVA03.[CANT_A_FAC_2]
										,GVA03.[CANT_PEDID_2]
										,GVA03.[CANT_PEN_D_2]
										,GVA03.[CANT_PEN_F_2]
										,GVA03.[PEN_REM_FC_2]
										,GVA03.[PEN_FAC_RE_2]
										,GVA03.[ID_MEDIDA_VENTAS]
										,GVA03.[ID_MEDIDA_STOCK_2]
										,GVA03.[ID_MEDIDA_STOCK]
										,GVA03.[UNIDAD_MEDIDA_SELECCIONADA]
										,GVA03.[COD_ARTICU_KIT]
										,GVA03.[RENGL_PADR]
										,GVA03.[PROMOCION]
										,GVA03.[PRECIO_ADICIONAL_KIT]
										,GVA03.[KIT_COMPLETO]
										,GVA03.[INSUMO_KIT_SEPARADO]
										,GVA03.[PRECIO_LISTA]
										,GVA03.[PRECIO_BONIF]
										,GVA03.[DESCUENTO_PARAM]
										,GVA03.[PRECIO_FECHA]
										,GVA03.[FECHA_MODIFICACION_PRECIO]
										,GVA03.[USUARIO_MODIFICACION_PRECIO]
										,GVA03.[TERMINAL_MODIFICACION_PRECIO]
										,GVA03.[ID_NEXO_PEDIDOS_RENGLON_ORDEN]
										,GVA03.[CANT_A_DES_EXPORTADA]
										,GVA03.[CANT_A_FAC_EXPORTADA]
										,GVA03.[CANT_A_DES_2_EXPORTADA]
										,GVA03.[CANT_A_FAC_2_EXPORTADA]
										,GVA03.[COD_DEPOSI]
									FROM [" + empresa_b + @"].[dbo].[GVA03]
									LEFT JOIN [" + empresa_b + @"]..GVA55 ON GVA55.TALON_PED = GVA03.TALON_PED AND GVA55.NRO_PEDIDO = GVA03.NRO_PEDIDO
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = gva55.T_COMP AND sta14_destino.N_COMP = gva55.N_COMP
									LEFT JOIN [" + empresa_a + @"]..GVA03 GVA03_destino ON 
										GVA03_destino.NRO_PEDIDO = GVA03.NRO_PEDIDO AND GVA03_destino.N_RENGLON = GVA03.N_RENGLON AND GVA03_destino.TALON_PED = GVA03.TALON_PED										
									WHERE GVA55.N_COMP = '" + n_comp + @"' AND GVA55.T_COMP = '" + t_comp + @"' AND GVA03_destino.NRO_PEDIDO IS NULL									
									";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar =
									@"INSERT INTO [" + empresa_a + @"].[dbo].[GVA45]
									([FILLER]
									,[COD_MODELO]
									,[DESC]
									,[DESC_ADIC]
									,[N_COMP]
									,[N_RENGLON]
									,[TALONARIO]
									,[T_COMP])
									SELECT 
										GVA45.[FILLER]
										,GVA45.[COD_MODELO]
										,GVA45.[DESC]
										,GVA45.[DESC_ADIC]
										,GVA45.[N_COMP]
										,GVA45.[N_RENGLON]
										,GVA45.[TALONARIO]
										,GVA45.[T_COMP]
									FROM [" + empresa_b + @"].[dbo].[GVA45]
									LEFT JOIN [" + empresa_b + @"]..GVA55 ON GVA55.TALON_PED = GVA45.TALONARIO AND GVA55.NRO_PEDIDO = GVA45.N_COMP
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = gva55.T_COMP AND sta14_destino.N_COMP = gva55.N_COMP
									LEFT JOIN [" + empresa_a + @"]..GVA45 GVA45_destino ON 
										GVA45_destino.N_COMP = GVA45.N_COMP AND GVA45_destino.N_RENGLON = GVA45.N_RENGLON
										AND GVA45_destino.T_COMP = GVA45.T_COMP AND GVA45_destino.TALONARIO = GVA45.TALONARIO										
									WHERE GVA55.N_COMP = '" + n_comp + @"' AND GVA55.T_COMP = '" + t_comp + @"' AND GVA45_destino.N_COMP IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									insertar = @"INSERT INTO [" + empresa_a + @"].[dbo].[GVA106]
									([FILLER]
									,[CANT_REM]
									,[FECHA_PED]
									,[FECHA_REM]
									,[NCOMP_IN_S]
									,[N_RENG_PED]
									,[N_RENGL_S]
									,[NRO_PEDIDO]
									,[TALON_PED]
									,[TCOMP_IN_S]
									,[CANTIDAD_REM_2]
									,[CANTIDAD_PEDIDO]
									,[CANTIDAD_PEDIDO_2])
									SELECT 
										GVA106.[FILLER]
										,GVA106.[CANT_REM]
										,GVA106.[FECHA_PED]
										,GVA106.[FECHA_REM]
										,sta14_destino.[NCOMP_IN_S]
										,GVA106.[N_RENG_PED]
										,GVA106.[N_RENGL_S]
										,GVA106.[NRO_PEDIDO]
										,GVA106.[TALON_PED]
										,GVA106.[TCOMP_IN_S]
										,GVA106.[CANTIDAD_REM_2]
										,GVA106.[CANTIDAD_PEDIDO]
										,GVA106.[CANTIDAD_PEDIDO_2]
									FROM [" + empresa_b + @"].[dbo].[GVA106] 									
									INNER JOIN [" + empresa_b + @"]..STA14 ON STA14.NCOMP_IN_S = GVA106.NCOMP_IN_S AND STA14.TCOMP_IN_S = GVA106.TCOMP_IN_S
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = sta14.T_COMP AND sta14_destino.N_COMP = sta14.N_COMP AND sta14_destino.talonario = sta14.talonario									
									LEFT JOIN [" + empresa_a + @"]..GVA106 GVA106_destino ON 
										GVA106_destino.NCOMP_IN_S = GVA106.NCOMP_IN_S AND GVA106_destino.TCOMP_IN_S = GVA106.TCOMP_IN_S
										AND GVA106_destino.NRO_PEDIDO = GVA106.NRO_PEDIDO AND GVA106_destino.TALON_PED = GVA106.TALON_PED
										AND GVA106_destino.N_RENG_PED = GVA106.N_RENG_PED AND GVA106_destino.N_RENGL_S = GVA106.N_RENGL_S
									WHERE GVA106.NCOMP_IN_S = '" + ncomp_in_s + @"' AND GVA106.TCOMP_IN_S = '" + tcomp_in_s + @"' AND GVA106_destino.NCOMP_IN_S IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();


									insertar = @"INSERT INTO [" + empresa_a + @"].[dbo].[GVA55]
									([FILLER]
									,[COD_CLIENT]
									,[FECHA_PEDI]
									,[N_COMP]
									,[NRO_PEDIDO]
									,[T_COMP]
									,[TALON_PED])
									SELECT 
										GVA55.[FILLER]
										,'"+cod_client+@"' [COD_CLIENT]
										,GVA55.[FECHA_PEDI]
										,GVA55.[N_COMP]
										,GVA55.[NRO_PEDIDO]
										,GVA55.[T_COMP]
										,GVA55.[TALON_PED]
									FROM [" + empresa_b + @"].[dbo].[GVA55] 
									INNER JOIN [" + empresa_a + @"]..sta14 sta14_destino ON sta14_destino.T_COMP = gva55.T_COMP AND sta14_destino.N_COMP = gva55.N_COMP
									LEFT JOIN [" + empresa_a + @"]..GVA55 GVA55_destino ON GVA55_destino.N_COMP = GVA55.N_COMP AND GVA55_destino.T_COMP = GVA55.T_COMP AND GVA55_destino.TALON_PED = GVA55.TALON_PED AND GVA55_destino.NRO_PEDIDO = GVA55.NRO_PEDIDO
									WHERE GVA55.N_COMP = '" + n_comp + @"' AND GVA55.T_COMP = '" + t_comp + @"' AND GVA55_destino.N_COMP IS NULL";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();

									Decimal saldo = 0;
									String cod_articu = "", cod_deposi = "", partida = "";
									sqlComando.CommandText =
									@"SELECT SUM((CASE WHEN STA20.TIPO_MOV = 'E' THEN 1 ELSE -1 END) * STA20.cantidad) SALDO, STA20.COD_ARTICU, STA20.COD_DEPOSI 
									FROM [" + empresa_a + @"]..STA20 
									WHERE STA20.COD_ARTICU IN (
										SELECT STA20.COD_ARTICU 
										FROM [" + empresa_b + @"]..STA20 
										WHERE STA20.NCOMP_IN_S = '" + ncomp_in_s + @"' AND STA20.TCOMP_IN_S = '" + tcomp_in_s + @"'
									)
									GROUP BY STA20.COD_ARTICU, STA20.COD_DEPOSI";
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();
									SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
									tablaConsulta2 = new DataTable();
									SqlAdaptadorDatos.Fill(tablaConsulta2);
									if (tablaConsulta2.Rows.Count > 0)
									{
										foreach (DataRow fila2 in tablaConsulta2.Rows)
										{
											saldo = Convert.ToDecimal(fila2["SALDO"]);
											cod_articu = fila2["COD_ARTICU"].ToString();
											cod_deposi = fila2["COD_DEPOSI"].ToString();

											insertar = @"UPDATE [" + empresa_a + @"].[dbo].[STA19] SET CANT_STOCK = '" + saldo.ToString().Replace(',', '.') + @"' WHERE cod_articu = '" + cod_articu + @"' AND cod_deposi = '" + cod_deposi + @"'";											
											sqlComando.CommandText = insertar;
											sqlComando.CommandType = CommandType.Text;
											sqlComando.ExecuteNonQuery();
										}
									}

									sqlComando.CommandText =
									@"SELECT SUM((CASE WHEN STA20.TIPO_MOV = 'E' THEN 1 ELSE -1 END) * STA09.cantidad) SALDO, STA09.COD_ARTICU, STA09.COD_DEPOSI, STA09.N_PARTIDA
									FROM [" + empresa_a + @"]..sta09 
									INNER JOIN [" + empresa_a + @"]..STA20 ON STA20.NCOMP_IN_S = STA09.NCOMP_IN_S and STA20.TCOMP_IN_S = STA09.TCOMP_IN_S and STA20.N_RENGL_S = STA09.N_RENGL_S
									WHERE STA09.COD_ARTICU IN (
										SELECT STA20.COD_ARTICU 
										FROM [" + empresa_b + @"]..STA20 
										WHERE STA20.NCOMP_IN_S = '" + ncomp_in_s + @"' AND STA20.TCOMP_IN_S = '" + tcomp_in_s + @"'
									)
									GROUP BY STA09.COD_ARTICU, STA09.COD_DEPOSI, STA09.N_PARTIDA";
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();
									SqlAdaptadorDatos = new SqlDataAdapter(sqlComando);
									tablaConsulta2 = new DataTable();
									SqlAdaptadorDatos.Fill(tablaConsulta2);
									if (tablaConsulta2.Rows.Count > 0)
									{
										foreach (DataRow fila2 in tablaConsulta2.Rows)
										{
											saldo = Convert.ToDecimal(fila2["SALDO"]);
											cod_articu = fila2["COD_ARTICU"].ToString();
											cod_deposi = fila2["COD_DEPOSI"].ToString();
											partida = fila2["N_PARTIDA"].ToString();

											insertar = @"UPDATE [" + empresa_a + @"].[dbo].[STA10] SET CANTIDAD = '" + saldo.ToString().Replace(',', '.') + @"' WHERE cod_articu = '" + cod_articu + @"' AND cod_deposi = '" + cod_deposi + @"' AND n_partida = '" + partida + @"'";
											sqlComando.CommandText = insertar;
											sqlComando.CommandType = CommandType.Text;
											sqlComando.ExecuteNonQuery();
										}
									}

								}
								else if(tiene_errores == "No")
								{
									generarLog("Hay errores en el remito " + n_comp);
									enviarMail(hayErrores, "Hay errores en el remito " + n_comp + "<br>");

									//insertar para que no vuelva a solicitar el error
									insertar = @"INSERT INTO [" + empresa_a + @"].[dbo].[STA14_ERRORES] ([TALONARIO],[N_COMP],[T_COMP],[FECHA]) VALUES ('" + talonario + @"', '" + n_comp + @"', '" + t_comp + @"', GETDATE())";
									sqlComando.CommandText = insertar;
									sqlComando.CommandType = CommandType.Text;
									sqlComando.ExecuteNonQuery();
								}
							}
						}
						else if(Convert.ToInt32(DateTime.Now.ToString("HH")) == 0 && Convert.ToInt32(DateTime.Now.ToString("mm")) >= 0 && Convert.ToInt32(DateTime.Now.ToString("mm")) <= 5 && se_borro_ultima_vez != DateTime.Now.ToString("yyyyMMdd"))
                        {
							//elimino para que se vuelva a intentar al principio del dia
							insertar = @"DELETE [" + empresa_a + @"].[dbo].[STA14_ERRORES]";
							sqlComando.CommandText = insertar;
							sqlComando.CommandType = CommandType.Text;
							sqlComando.ExecuteNonQuery();

							se_borro_ultima_vez = DateTime.Now.ToString("yyyyMMdd");
                        }

						StreamWriter sw = new StreamWriter("ultima_ejecucion.txt", false);
						sw.WriteLine(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff") +" | " + sql);
						sw.Close();

						transaction.Commit();
					}
                    catch (SqlException ex)
                    {
                        transaction.Rollback();

                        generarLog(ex.ToString());
                        hayErrores.Add("Error de conexión, comuniquese con su administrador de base de datos");
                        enviarMail(hayErrores, "Hay errores de conexion: <br>");
						/*
                        Timer_informacion_error();
                        timer_informacion.Stop();
                        timer_informacion.Dispose();
						*/
                    }
                    finally
                    {
                        if (con.State == ConnectionState.Open)
                            con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                generarLog(ex.ToString());
                hayErrores.Add("Error en el codigo");
                enviarMail(hayErrores, "Hay errores generales: <br>");
            }

        }

        private static void generarLog(string error)
        {
            DateTime dateTime = DateTime.Now;
            string strDate = Convert.ToDateTime(dateTime).ToString("dd/MM/yyyy HH:mm:ss.fff");
            using (StreamWriter archivo = File.AppendText(@"log.txt"))
            {
                archivo.WriteLine(strDate + " - " + error);
                archivo.Close();
            }
        }

        private static void enviarMail(List<String> hayErrores, String contenido_mail)
        {
			try
			{
				MailMessage mail;
				mail = new MailMessage();
				mail.From = new MailAddress("crmflow2020@gmail.com");

				string[] correosElectronicos = emailDestinatarios.Split(';');
				foreach (var correoElectronico in correosElectronicos)
				{
					mail.To.Add(new MailAddress(correoElectronico));
				}

				/*
				string[] rutasArchivos = rutaArchivos.Split(';');
				foreach (var file in rutasArchivos)
				{
					Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);
					ContentDisposition disposition = data.ContentDisposition;
					disposition.CreationDate = System.IO.File.GetCreationTime(file);
					disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
					disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
					mail.Attachments.Add(data);
				}
				*/

				foreach (var error in hayErrores)
				{
					contenido_mail += error + "<br>";
				}
				contenido_mail += "<br><br>Saludos";

				mail.Subject = "Tango - Error al sincronizar remito";
				mail.Body = contenido_mail;
				mail.IsBodyHtml = true;
				SmtpClient client = new SmtpClient("smtp.gmail.com", 25);
				using (client)
				{
					client.Credentials = new System.Net.NetworkCredential("crmflow2020@gmail.com", "hrcwgpdrqznaieiu");
					client.EnableSsl = true;
					client.Send(mail);
				}
			}catch(Exception ex)
            {
				generarLog(ex.Message);
            }
        }

    }
}
