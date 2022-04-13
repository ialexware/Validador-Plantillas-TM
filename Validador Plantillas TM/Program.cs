using ConsoleTables;
using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Validador_Plantillas_TM.Clases;
using Validador_Plantillas_TM.Enumerables;
using Validador_Plantillas_TM.Modelos;

namespace Validador_Plantillas_TM
{
    abstract class Validador
    {

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int cmdShow);

        private static List<ClaveProdServCP> lstClaveProdServCP;
        private static List<ClaveUnidad> lstCatClaveUnidad;
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            SetFullsize();

            //string strFileName = $"Layout_CartaPorte Clientes MVL_1";
            string strFileName = $"Layout_CartaPorte Clientes MVL_1 - copia";
            //string strFileName = $"Layuot para ccp  22328  2227";
            //string strFileName = $"Layuot para ccp  22328  22271Hoja";

            //string strFileName = $"LayoutCP2";
            //string strFileName = $"LayoutCP3";
            //string strFileName = $"CS1017 BOMBARDIER";
            //string strFileName = $"CS1017 BOMBARDIERPedimento"; // Error pedimento, Sí No
            //string strFileName = $"Carta Porte 31625_31_mar_2022";
            string strFileFullName = $"{strFileName}{".xlsx"}";

            string strBaseDir = Directory.GetCurrentDirectory();
            Console.WriteLine(strBaseDir);

            string strDirectorioCatalogos = Path.Combine(strBaseDir, "Catalogos");
            string strCatClaveProdServCPDir = Path.Combine(strBaseDir, "Catalogos", "c_ClaveProdServCP.json");
            string strCatClaveUnidadDir = Path.Combine(strBaseDir, "Catalogos", "c_ClaveUnidad.json");

            string strDirectorioEntrada = Path.Combine(strBaseDir, "Entrada");
            string strDirectorioSalida = Path.Combine(strBaseDir, "Salida");
            string strDirectorioError = Path.Combine(strBaseDir, "Error");



            //  [08 / 04 / 2022 03:03 p.m.] Raul Santillan Pulido
            //  1 ) La cantidad no debe ser 0, máximo de 6 posiciones decimales valor mínimo 0.000001

            //  [08 / 04 / 2022 03:05 p.m.] Raul Santillan Pulido
            //  2 ) Revisar material peligroso
            //  Valores permitidos; Si, No(Cualquier otro caso mandar error)
            //  Validar de acuerdo la claveProdServ en la columna de MaterialPeligroso
            //  Cuando sea un Si debe contener:
            //  CveMaterialPeligroso - Validar de catálogo c_MaterialPeligroso
            //  Embalaje - Validar de c_TipoEmbalaje
            //  Descripción Embalaje -0 - 100 caracteres

            //  [08 / 04 / 2022 03:06 p.m.] Raul Santillan Pulido
            //  3 ) PesoenKg no debe ser cero, valor mínimo 0.001

            //  [08 / 04 / 2022 03:06 p.m.] Raul Santillan Pulido
            //  4 ) El nombre de las columnas debe ser tal como viene en el layout


            //  Si existe Importación , Validar pedimento Correcto en estructura


            //  Bandera de Transporte Internacional si es Transporte internacional especifica si es Importación o Exportación
            //  Ambos casos Necesita FraccionArancelaria,
            //  Exportación necesita UUIDComercioExt - Patron en documento adjunto
            //  Importación necesita Pedimento, Cuando exista más de un pedimento están separadas por comas.Patron en documento adjunto
            bool blnInternacional = false;
            TipoComercioPlantilla tipoComercioPlantilla = TipoComercioPlantilla.Exportacion;


            try


            {


                // Si el directorio no existe entonces se crea
                if (
                //CheckTargetDirectory(strDirectorioCatalogos) &&
                CheckTargetDirectory(strDirectorioEntrada) &&
                CheckTargetDirectory(strDirectorioSalida) &&
                CheckTargetDirectory(strDirectorioError))
                {
                    Console.WriteLine("Directorios validados ");
                    //Lectura de los catálogos 
                    bool blnContinuar = LecturaCatalogos(strCatClaveProdServCPDir, strCatClaveUnidadDir);

                    if (blnContinuar)
                    {
                        Console.WriteLine("Lectura de catálogos correcta");

                        var strFullPath = Path.Combine(strDirectorioEntrada, strFileFullName);
                        RespuestaProcesoArchivo objRespuesta = LecturaDePlantilla(strFullPath, blnInternacional, tipoComercioPlantilla);
                        if (objRespuesta != null)
                        {

                            MostrarResultado(objRespuesta, strFileFullName, strDirectorioError, strFileName);
                        }
                        else
                        {
                            Console.WriteLine("Algo salió mal al leer la plantilla.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Algo salió mal al leer los catálogos.");
                    }
                }
                else
                {
                    Console.WriteLine("Algo salió mal al crear los directorios.");
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
            }
            Console.ReadKey();
        }

        private static RespuestaProcesoArchivo LecturaDePlantilla(string strFullPath, bool blnInternacional, TipoComercioPlantilla tipoComercioPlantilla)
        {
            RespuestaProcesoArchivo objRespuesta = new RespuestaProcesoArchivo();
            try
            {
                List<ModeloArchivo> lstBaseAplicar = new List<ModeloArchivo>();



                var intIteracion = 1;

                using (var progress = new ProgressBar())
                {
                    //var pb = new Util.ProgressBar("Analyzing data");
                    //pb.Dump();



                    using (var stream = File.Open(strFullPath, FileMode.Open, FileAccess.Read))
                    {



                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            objRespuesta.intTotalFilas = reader.RowCount;
                            objRespuesta.intTotalHojas = reader.ResultsCount;
                            objRespuesta.intTotalColumnas = reader.FieldCount;
                            objRespuesta.strNombreHoja = reader.Name;

                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true,
                                },
                                UseColumnDataType = true
                            });


                            var columns = result.Tables[0].Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();

                            int intIdexColumnOrigen = columns.IndexOf(columns.Where(c => c.Contains("Origen")).FirstOrDefault());
                            int intIdexColumnDestinoFinal = columns.IndexOf(columns.Where(c => c.Contains("Destino final")).FirstOrDefault());
                            int intIdexColumnBienesTransp = columns.IndexOf(columns.Where(c => c.Contains("BienesTransp")).FirstOrDefault());
                            int intIdexColumnDescripcion = columns.IndexOf(columns.Where(c => c.Contains("Descripcion")).FirstOrDefault());
                            int intIdexColumnCantidad = columns.IndexOf(columns.Where(c => c.Contains("Cantidad")).FirstOrDefault());
                            int intIdexColumnClaveUnidad = columns.IndexOf(columns.Where(c => c.Contains("ClaveUnidad")).FirstOrDefault());
                            int intIdexColumnUnidad = columns.IndexOf(columns.Where(c => c == "Unidad").FirstOrDefault());
                            int intIdexColumnMaterialPeligroso = columns.IndexOf(columns.Where(c => c.Contains("MaterialPeligroso")).FirstOrDefault());
                            int intIdexColumnCveMaterialPeligroso = columns.IndexOf(columns.Where(c => c.Contains("CveMaterialPeligroso")).FirstOrDefault());
                            int intIdexColumnEmbalaje = columns.IndexOf(columns.Where(c => c.Contains("Embalaje")).FirstOrDefault());
                            int intIdexColumnDescripEmbalaje = columns.IndexOf(columns.Where(c => c.Contains("DescripEmbalaje")).FirstOrDefault());
                            int intIdexColumnPesoEnKg = columns.IndexOf(columns.Where(c => c.Contains("PesoEnKg")).FirstOrDefault());
                            int intIdexColumnFraccionArancelaria = columns.IndexOf(columns.Where(c => c.Contains("FraccionArancelaria")).FirstOrDefault());
                            int intIdexColumnUUIDComercioExt = columns.IndexOf(columns.Where(c => c.Contains("UUIDComercioExt")).FirstOrDefault());
                            int intIdexColumnPedimento = columns.IndexOf(columns.Where(c => c.Contains("Pedimento")).FirstOrDefault());
                            int intIdexColumnRepartos = columns.IndexOf(columns.Where(c => c.Contains("Repartos")).FirstOrDefault());
                            int intIdexColumnGuia = columns.IndexOf(columns.Where(c => c.Contains("Guia")).FirstOrDefault());



                            do
                            {
                                while (reader.Read())
                                {
                                    List<Incidencia> lstIncLinea = new List<Incidencia>();

                                    progress.Report(((intIteracion * 100) / objRespuesta.intTotalFilas));


                                    ModeloArchivo itmPrimeraFila = new ModeloArchivo();
                                    itmPrimeraFila.intFila = intIteracion;

                                    if (reader.Name.ToUpper() == "LAYOUT")
                                    {

                                        if (intIdexColumnOrigen >= 0)
                                        {
                                            itmPrimeraFila.Origen = Convert.ToString(reader.GetValue(intIdexColumnOrigen)).Trim();
                                        }
                                        if (intIdexColumnDestinoFinal >= 0)
                                        {
                                            itmPrimeraFila.DestinoFinal = Convert.ToString(reader.GetValue(intIdexColumnDestinoFinal)).Trim();
                                        }
                                        if (intIdexColumnBienesTransp >= 0)
                                        {
                                            itmPrimeraFila.BienesTransp = Convert.ToString(reader.GetValue(intIdexColumnBienesTransp)).Trim();
                                        }
                                        if (intIdexColumnDescripcion >= 0)
                                        {
                                            itmPrimeraFila.Descripcion = Convert.ToString(reader.GetValue(intIdexColumnDescripcion)).Trim();
                                        }
                                        if (intIdexColumnCantidad >= 0)
                                        {
                                            itmPrimeraFila.Cantidad = Convert.ToString(reader.GetValue(intIdexColumnCantidad)).Trim();
                                        }
                                        if (intIdexColumnClaveUnidad >= 0)
                                        {
                                            itmPrimeraFila.ClaveUnidad = Convert.ToString(reader.GetValue(intIdexColumnClaveUnidad)).Trim();
                                        }
                                        if (intIdexColumnUnidad >= 0)
                                        {
                                            itmPrimeraFila.Unidad = Convert.ToString(reader.GetValue(intIdexColumnUnidad)).Trim();
                                        }
                                        if (intIdexColumnMaterialPeligroso >= 0)
                                        {
                                            itmPrimeraFila.MaterialPeligroso = Convert.ToString(reader.GetValue(intIdexColumnMaterialPeligroso)).Trim();
                                        }
                                        if (intIdexColumnCveMaterialPeligroso >= 0)
                                        {
                                            itmPrimeraFila.CveMaterialPeligroso = Convert.ToString(reader.GetValue(intIdexColumnCveMaterialPeligroso)).Trim();
                                        }
                                        if (intIdexColumnEmbalaje >= 0)
                                        {
                                            itmPrimeraFila.Embalaje = Convert.ToString(reader.GetValue(intIdexColumnEmbalaje)).Trim();
                                        }
                                        if (intIdexColumnDescripEmbalaje >= 0)
                                        {
                                            itmPrimeraFila.DescripEmbalaje = Convert.ToString(reader.GetValue(intIdexColumnDescripEmbalaje)).Trim();
                                        }
                                        if (intIdexColumnPesoEnKg >= 0)
                                        {
                                            itmPrimeraFila.PesoEnKg = Convert.ToString(reader.GetValue(intIdexColumnPesoEnKg)).Trim();
                                        }
                                        if (intIdexColumnFraccionArancelaria >= 0)
                                        {
                                            itmPrimeraFila.FraccionArancelaria = Convert.ToString(reader.GetValue(intIdexColumnFraccionArancelaria)).Trim();
                                        }
                                        if (intIdexColumnUUIDComercioExt >= 0)
                                        {
                                            itmPrimeraFila.UUIDComercioExt = Convert.ToString(reader.GetValue(intIdexColumnUUIDComercioExt)).Trim();
                                        }
                                        if (intIdexColumnPedimento >= 0)
                                        {
                                            itmPrimeraFila.Pedimento = Convert.ToString(reader.GetValue(intIdexColumnPedimento)).Trim();
                                        }



                                        // Cabecero 
                                        if (intIteracion == 1)
                                        {
                                            if (objRespuesta.intTotalColumnas < 15)
                                            {
                                                Incidencia itmIncidencia = new Incidencia()
                                                {
                                                    Linea = 0,
                                                    Columna = "",
                                                    Valor = objRespuesta.strNombreHoja,
                                                    ValorEsperado = "‘Layout’",
                                                    Mensaje = "El nombre de la hoja debe ser ‘Layout’"
                                                };
                                                lstIncLinea.Add(itmIncidencia);
                                            }

                                            if (objRespuesta.intTotalColumnas < 15)
                                            {
                                                Incidencia itmIncidencia = new Incidencia()
                                                {
                                                    Linea = intIteracion,
                                                    Columna = "",
                                                    Valor = objRespuesta.intTotalColumnas.ToString(),
                                                    ValorEsperado = "15",
                                                    Mensaje = "Se esperaban 15 columnas"
                                                };
                                                lstIncLinea.Add(itmIncidencia);
                                            }
                                            else if (objRespuesta.intTotalColumnas > 15)
                                            {

                                                if (intIdexColumnGuia > 0)
                                                {
                                                    itmPrimeraFila.Guia = Convert.ToString(reader.GetValue(intIdexColumnGuia)).Trim();
                                                }

                                                if (intIdexColumnRepartos > 0)
                                                {
                                                    itmPrimeraFila.Repartos = Convert.ToString(reader.GetValue(intIdexColumnRepartos)).Trim();
                                                }

                                            }

                                            if (intIdexColumnOrigen >= 0)
                                            {
                                                //  Origen
                                                ValidaHeaderColumna(itmPrimeraFila.Origen, "Origen", "Origen", intIteracion, ref lstIncLinea);
                                            }

                                            //  Destino final 
                                            if (intIdexColumnDestinoFinal >= 0)
                                            {
                                                ValidaHeaderColumna(itmPrimeraFila.DestinoFinal, "Destino final", "Destino final", intIteracion, ref lstIncLinea);
                                            }

                                            //  BienesTransp
                                            ValidaHeaderColumna(itmPrimeraFila.BienesTransp, "BienesTransp", "BienesTransp", intIteracion, ref lstIncLinea);

                                            //  Descripcion
                                            ValidaHeaderColumna(itmPrimeraFila.Descripcion, "Descripcion", "Descripcion", intIteracion, ref lstIncLinea);

                                            //  Cantidad
                                            ValidaHeaderColumna(itmPrimeraFila.Cantidad, "Cantidad", "Cantidad", intIteracion, ref lstIncLinea);

                                            //  ClaveUnidad
                                            ValidaHeaderColumna(itmPrimeraFila.ClaveUnidad, "ClaveUnidad", "ClaveUnidad", intIteracion, ref lstIncLinea);

                                            //  Unidad
                                            ValidaHeaderColumna(itmPrimeraFila.Unidad, "Unidad", "Unidad", intIteracion, ref lstIncLinea);

                                            //  MaterialPeligroso
                                            ValidaHeaderColumna(itmPrimeraFila.MaterialPeligroso, "MaterialPeligroso", "MaterialPeligroso", intIteracion, ref lstIncLinea);

                                            //  CveMaterialPeligroso
                                            ValidaHeaderColumna(itmPrimeraFila.CveMaterialPeligroso, "CveMaterialPeligroso", "CveMaterialPeligroso", intIteracion, ref lstIncLinea);

                                            //  Embalaje
                                            ValidaHeaderColumna(itmPrimeraFila.Embalaje, "Embalaje", "Embalaje", intIteracion, ref lstIncLinea);

                                            //  DescripEmbalaje
                                            ValidaHeaderColumna(itmPrimeraFila.DescripEmbalaje, "DescripEmbalaje", "DescripEmbalaje", intIteracion, ref lstIncLinea);

                                            //  PesoEnKg
                                            ValidaHeaderColumna(itmPrimeraFila.PesoEnKg, "PesoEnKg", "PesoEnKg", intIteracion, ref lstIncLinea);

                                            //  FraccionArancelaria
                                            ValidaHeaderColumna(itmPrimeraFila.FraccionArancelaria, "FraccionArancelaria", "FraccionArancelaria", intIteracion, ref lstIncLinea);

                                            //  UUIDComercioExt
                                            ValidaHeaderColumna(itmPrimeraFila.UUIDComercioExt, "UUIDComercioExt", "UUIDComercioExt", intIteracion, ref lstIncLinea);

                                            //  Pedimento
                                            ValidaHeaderColumna(itmPrimeraFila.Pedimento, "Pedimento", "Pedimento", intIteracion, ref lstIncLinea);

                                            if (intIdexColumnRepartos > 0)
                                            {
                                                //  Repartos
                                                ValidaHeaderColumna(itmPrimeraFila.Guia, "Repartos", "Repartos", intIteracion, ref lstIncLinea);
                                            }

                                            if (intIdexColumnGuia > 0)
                                            {
                                                //  Guia
                                                ValidaHeaderColumna(itmPrimeraFila.Guia, "Guia", "Guia", intIteracion, ref lstIncLinea);
                                            }


                                        }
                                        //  Cuerpo del archivo
                                        else if (intIteracion > 1)
                                        {

                                            if (intIdexColumnOrigen >= 0)
                                            {
                                                //  Origen
                                                CeldaEnBlanco(itmPrimeraFila.Origen, "Origen de la mercancía.", "Origen", intIteracion, ref lstIncLinea);
                                            }

                                            if (intIdexColumnDestinoFinal >= 0)
                                            {
                                                //  Destino final 
                                                CeldaEnBlanco(itmPrimeraFila.DestinoFinal, "Destino de la mercancía.", "Destino final", intIteracion, ref lstIncLinea);
                                            }

                                            //  BienesTransp
                                            CeldaEnBlanco(itmPrimeraFila.BienesTransp, "BienesTransp", "BienesTransp", intIteracion, ref lstIncLinea);
                                            ClaveProdServCP claveProdServCP = BuscarBienesTranspCatalogo(itmPrimeraFila.BienesTransp, "BienesTransp", intIteracion, ref lstIncLinea);

                                            //  Descripcion
                                            CeldaEnBlanco(itmPrimeraFila.Descripcion, "Atributo requerido para detallar las características de los bienes y / o mercancías.", "Descripcion", intIteracion, ref lstIncLinea);

                                            //  Cantidad
                                            CeldaEnBlanco(itmPrimeraFila.Cantidad, "Atributo requerido para expresar la cantidad total de los bienes.  > 0.000001", "Cantidad", intIteracion, ref lstIncLinea);
                                            CeldaDecimalValorMinimo(itmPrimeraFila.Cantidad, 0.000001, "Cantidad", intIteracion, ref lstIncLinea);


                                            //  ClaveUnidad
                                            CeldaEnBlanco(itmPrimeraFila.ClaveUnidad, "Atributo requerido para registrar la clave de la unidad de medida  estandarizada aplicable para la  cantidad de los bienes.", "ClaveUnidad", intIteracion, ref lstIncLinea);
                                            ClaveUnidad claveUnidad = BuscarUnidadCatalogo(itmPrimeraFila.ClaveUnidad, "ClaveUnidad", intIteracion, ref lstIncLinea);

                                            //  Unidad
                                            //CeldaEnBlanco(itmPrimeraFila.Unidad, "Unidad de la mercancía.", "Unidad", intIteracion, ref lstIncLinea);



                                            //  MaterialPeligroso
                                            ValidaMaterialPeligroso(itmPrimeraFila.MaterialPeligroso, "MaterialPeligroso", intIteracion, ref lstIncLinea, ref itmPrimeraFila);
                                            bool blnMtPel = ValidaMaterialPeligrosoCatalogo("MaterialPeligroso", intIteracion, itmPrimeraFila.BienesTransp, ref lstIncLinea, itmPrimeraFila, claveProdServCP);

                                            if (claveProdServCP != null && (claveProdServCP.blnMaterialPeligroso || blnMtPel))
                                            {
                                                //  CveMaterialPeligroso
                                                CeldaEnBlancoMaterial(itmPrimeraFila.CveMaterialPeligroso, "CveMaterialPeligroso", "CveMaterialPeligroso", intIteracion, itmPrimeraFila.BienesTransp, ref lstIncLinea);

                                                //  Embalaje
                                                CeldaEnBlancoMaterial(itmPrimeraFila.Embalaje, "Embalaje", "Embalaje", intIteracion, itmPrimeraFila.BienesTransp, ref lstIncLinea);

                                                //  DescripEmbalaje
                                                CeldaEnBlancoMaterial(itmPrimeraFila.DescripEmbalaje, "DescripEmbalaje", "DescripEmbalaje", intIteracion, itmPrimeraFila.BienesTransp, ref lstIncLinea);
                                            }

                                            //  PesoEnKg
                                            CeldaEnBlanco(itmPrimeraFila.PesoEnKg, "PesoEnKg", "PesoEnKg", intIteracion, ref lstIncLinea);
                                            CeldaDecimalValorMinimo(itmPrimeraFila.PesoEnKg, 0.001, "PesoEnKg", intIteracion, ref lstIncLinea);







                                            if (blnInternacional)
                                            {
                                                //  FraccionArancelaria
                                                CeldaEnBlanco(itmPrimeraFila.FraccionArancelaria, "FraccionArancelaria", "FraccionArancelaria", intIteracion, ref lstIncLinea);

                                                //  UUIDComercioExt
                                                CeldaEnBlanco(itmPrimeraFila.UUIDComercioExt, "UUIDComercioExt", "UUIDComercioExt", intIteracion, ref lstIncLinea);

                                                //  Pedimento
                                                CeldaEnBlanco(itmPrimeraFila.Pedimento, "Pedimento", "Pedimento", intIteracion, ref lstIncLinea);

                                                if (tipoComercioPlantilla == TipoComercioPlantilla.Importacion)
                                                {
                                                    ValidaPedimentoCorrecto(itmPrimeraFila.Pedimento, "Pedimento", intIteracion, ref lstIncLinea);
                                                }

                                            }
                                            else
                                            {
                                                string strMensajeVacio = "La celda debe de estar vacía debido a que no se trata de una Importación/Exportación";
                                                //  FraccionArancelaria
                                                CeldaConValor(itmPrimeraFila.FraccionArancelaria, strMensajeVacio, "FraccionArancelaria", intIteracion, ref lstIncLinea);

                                                //  UUIDComercioExt
                                                CeldaConValor(itmPrimeraFila.UUIDComercioExt, strMensajeVacio, "UUIDComercioExt", intIteracion, ref lstIncLinea);

                                                //  Pedimento
                                                CeldaConValor(itmPrimeraFila.Pedimento, strMensajeVacio, "Pedimento", intIteracion, ref lstIncLinea);
                                            }




                                            if (intIdexColumnRepartos > 0)
                                            {
                                                //  Repartos
                                                CeldaEnBlanco(itmPrimeraFila.Guia, "Repartos", "Repartos", intIteracion, ref lstIncLinea);
                                            }

                                            if (intIdexColumnGuia > 0)
                                            {
                                                //  Guia
                                                //CeldaEnBlanco(itmPrimeraFila.Guia, "Guia", "Guia", intIteracion, ref lstIncLinea);
                                            }

                                        }
                                    }
                                    else
                                    {

                                        lstIncLinea.Add(new Incidencia()
                                        {
                                            Linea = intIteracion,
                                            Columna = "",
                                            Valor = reader.Name,
                                            ValorEsperado = "Solo 1 Hoja",
                                            Mensaje = $"'Existen datos inválidos en la Hoja { reader.Name }."
                                        });

                                    }
                                    //  Validamos la existencia de algún error en la línea actual 
                                    if (!lstIncLinea.Any())
                                    {
                                        lstBaseAplicar.Add(itmPrimeraFila);
                                        objRespuesta.intTotalFilasOk++;
                                    }
                                    else
                                    {
                                        objRespuesta.intTotalFilasError++;
                                    }

                                    //	Agregamos los errores de la línea actual al resumen general
                                    objRespuesta.lstIncidencias.AddRange(lstIncLinea);
                                    //	Aumentamos el contador la línea actual 
                                    intIteracion++;

                                }

                            } while (reader.NextResult());

                        }
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return objRespuesta;
        }

        private static void ValidaMaterialPeligroso(string strValor, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea, ref ModeloArchivo itmPrimeraFila)
        {
            List<string> strValoresEsperados = new List<string> { "Sí", "No" }.ToList();
            if (strValoresEsperados.Contains(strValor))
            {
                switch (strValor)
                {
                    case "Sí":
                        itmPrimeraFila.blnMaterialPeligroso = true;
                        break;
                    case "No":
                        itmPrimeraFila.blnMaterialPeligroso = false;
                        break;
                }
            }
            else
            {
                lstIncidenciasLinea.Add(new Incidencia()
                {
                    Linea = intIteracion,
                    Columna = strColumna,
                    Valor = (strValor != null) ? (strValor.Length > 10) ? strValor.Substring(0, 10) : strValor : "",
                    ValorEsperado = "'Sí', 'No'",
                    Mensaje = "'MaterialPeligroso' Los valores esperados en esta columna son: ‘Sí’ o ‘No’ Revisar catálogos en ‘http://omawww.sat.gob.mx/tramitesyservicios/Paginas/complemento_carta_porte.htm’   Catálogo de productos y servicios carta porte"
                });
            }
        }

        private static bool ValidaMaterialPeligrosoCatalogo(string strColumna, int intIteracion, string strMaterial, ref List<Incidencia> lstIncidenciasLinea, ModeloArchivo itmPrimeraFila, ClaveProdServCP claveProdServCP)
        {
            try
            {
                if (claveProdServCP != null && itmPrimeraFila != null)
                {
                    if (claveProdServCP.MaterialPeligroso != "0,1" && claveProdServCP.blnMaterialPeligroso != itmPrimeraFila.blnMaterialPeligroso)
                    {
                        lstIncidenciasLinea.Add(new Incidencia()
                        {
                            Linea = intIteracion,
                            Columna = strColumna,
                            Valor = $"{ itmPrimeraFila.blnMaterialPeligroso }",
                            ValorEsperado = $"{ claveProdServCP.blnMaterialPeligroso }",
                            Mensaje = $"Por favor valide el valor de ‘MaterialPeligroso’ para la clave { strMaterial } el valor proporcionado es { itmPrimeraFila.blnMaterialPeligroso } mientras que el catalogo dice { claveProdServCP.blnMaterialPeligroso } "
                        });
                    }
                    else if (claveProdServCP.MaterialPeligroso == "0,1" && itmPrimeraFila.blnMaterialPeligroso == true)
                    {
                        return true;
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return claveProdServCP.blnMaterialPeligroso;
        }

        private static void ValidaHeaderColumna(string strValor, string strValorEsperado, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (string.IsNullOrEmpty(strValor))
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = strValor,
                        ValorEsperado = strValorEsperado,
                        Mensaje = "No encontramos la columna"
                    });
                }
                else if (strValor.Contains(" ") && strValor != "Destino final")
                {
                    var lstValores = strValor.Split(' ');

                    if (lstValores[0] != strValorEsperado)
                    {
                        lstIncidenciasLinea.Add(new Incidencia()
                        {
                            Linea = intIteracion,
                            Columna = strColumna,
                            Valor = strValor,
                            ValorEsperado = strValorEsperado,
                            Mensaje = "No encontramos la columna"
                        });
                    }

                }
                else if (strValor != strValorEsperado)
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = strValor,
                        ValorEsperado = strValorEsperado,
                        Mensaje = "No encontramos la columna"
                    });
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static void ValidaPedimentoCorrecto(string strValor, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                string pattern = "[0-9]{2} [0-9]{2} [0-9]{4} [0-9]{7}";
                int intMaxSegundosTimeOut = 20;

                Regex regx = new Regex(pattern, RegexOptions.Singleline, new TimeSpan(0, 0, intMaxSegundosTimeOut));
                bool blnCorrecto = regx.IsMatch(strValor);
                if (!blnCorrecto)
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = strValor,
                        ValorEsperado = pattern,
                        Mensaje = $"El pedimento no cumple con el formato requerido: { pattern }"
                    });
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static void CeldaEnBlanco(string strValor, string strValorEsperado, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (strValor == null || string.IsNullOrEmpty(strValor))
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = (strValor != null) ? (strValor.Length > 10) ? strValor.Substring(0, 10) : strValor : "",
                        ValorEsperado = strValorEsperado,
                        Mensaje = $"'{strColumna}' CELDA EN BLANCO"
                    });
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static void CeldaEnBlancoMaterial(string strValor, string strValorEsperado, string strColumna, int intIteracion, string strMaterial, ref List<Incidencia> lstIncidenciasLinea)
        {

            try
            {
                if (strValor == null || string.IsNullOrEmpty(strValor))
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = (strValor != null) ? (strValor.Length > 10) ? strValor.Substring(0, 10) : strValor : "",
                        ValorEsperado = strValorEsperado,
                        Mensaje = $"'Celda '{ strColumna }' en blanco y es un valor requerido para el material peligroso  { strMaterial }"
                    });
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static void CeldaConValor(string strValor, string strValorEsperado, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (!string.IsNullOrEmpty(strValor))
                {
                    lstIncidenciasLinea.Add(new Incidencia()
                    {
                        Linea = intIteracion,
                        Columna = strColumna,
                        Valor = (strValor != null) ? (strValor.Length > 10) ? strValor.Substring(0, 10) : strValor : "",
                        ValorEsperado = strValorEsperado,
                        Mensaje = $"'{strColumna}' CELDA CON VALORES"
                    });
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static void CeldaDecimalValorMinimo(string strValor, double dblValorEsperado, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (!string.IsNullOrEmpty(strValor))
                {

                    bool success = double.TryParse(strValor, out double dblValor);
                    if (success)
                    {

                        if (dblValor <= 0)
                        {
                            lstIncidenciasLinea.Add(new Incidencia()
                            {
                                Linea = intIteracion,
                                Columna = strColumna,
                                Valor = $"{ dblValor:N10}",
                                ValorEsperado = $"Valor Mayor a { dblValorEsperado:N6}",
                                Mensaje = $"El valor de la celda debe ser mayor a  {dblValorEsperado:N6}"
                            });
                        }
                        else if (dblValor < dblValorEsperado)
                        {
                            lstIncidenciasLinea.Add(new Incidencia()
                            {
                                Linea = intIteracion,
                                Columna = strColumna,
                                Valor = $"{ dblValor:N10}",
                                ValorEsperado = $"Valor Mayor a { dblValorEsperado:N6}",
                                Mensaje = $"El valor de la celda debe ser mayor a  {dblValorEsperado:N6}"
                            });
                        }
                    }
                    else
                    {

                        lstIncidenciasLinea.Add(new Incidencia()
                        {
                            Linea = intIteracion,
                            Columna = strColumna,
                            Valor = strValor,
                            ValorEsperado = $"Valor Mayor a { dblValorEsperado:N6}",
                            Mensaje = $"El valor de la celda debe ser mayor a  { dblValorEsperado:N6}"
                        });
                    }


                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
        }

        private static ClaveProdServCP BuscarBienesTranspCatalogo(string strValor, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (!string.IsNullOrEmpty(strValor))
                {
                    var itCat = lstClaveProdServCP.Where(a => a.CClaveProdServ == strValor).FirstOrDefault();
                    if (itCat == null)
                    {
                        lstIncidenciasLinea.Add(new Incidencia()
                        {
                            Linea = intIteracion,
                            Columna = strColumna,
                            Valor = strValor,
                            ValorEsperado = strColumna,
                            Mensaje = $"El código {strValor} no se encontró dentro del catalogo c_ClaveProdServCP"
                        });
                    }
                    else
                    {
                        return itCat;
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return null;
        }

        private static ClaveUnidad BuscarUnidadCatalogo(string strValor, string strColumna, int intIteracion, ref List<Incidencia> lstIncidenciasLinea)
        {
            try
            {
                if (!string.IsNullOrEmpty(strValor))
                {
                    var itCat = lstCatClaveUnidad.Where(a => a.CClaveUnidad == strValor).FirstOrDefault();
                    if (itCat == null)
                    {
                        lstIncidenciasLinea.Add(new Incidencia()
                        {
                            Linea = intIteracion,
                            Columna = strColumna,
                            Valor = strValor,
                            ValorEsperado = strColumna,
                            Mensaje = $"Launidad {strValor} no se encontró dentro del catalogo c_ClaveUnidad"
                        });
                    }
                    else
                    {
                        return itCat;
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return null;
        }

        private static void MostrarResultado(RespuestaProcesoArchivo objRespuesta, string strFileFullName, string filePath, string strFileName)
        {
            List<string> lstMensajesFinales = new List<string>();

            if (objRespuesta.lstIncidencias.Any())
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
            }

            lstMensajesFinales.Add($"Archivo: {strFileFullName}");
            lstMensajesFinales.Add($"Total de registros leídos: { string.Format("{0:n0}", objRespuesta.intTotalFilas) }");
            lstMensajesFinales.Add($"Total de registros con Error: { string.Format("{0:n0}", objRespuesta.intTotalFilasError)}");
            lstMensajesFinales.Add($"Total de registros con ok: { string.Format("{0:n0}", objRespuesta.intTotalFilasOk)}");
            lstMensajesFinales.Add($"Total de Errores: { string.Format("{0:n0}", objRespuesta.lstIncidencias.Count())}");

            if (objRespuesta.lstIncidencias.Any())
            {
                lstMensajesFinales.Add("Resultado: El archivo Proporcionado NO cumple de forma general con la estructura solicitada ");
            }
            else
            {
                lstMensajesFinales.Add("Resultado: El archivo Proporcionado cumple de forma general con la estructura solicitada ");
            }

            //var table = new ConsoleTable(objRespuesta.lstIncidencias);
            var table = new ConsoleTable("Mensajes:");
            lstMensajesFinales.ForEach(f => { table.AddRow(f); });
            table.Write();

            List<Incidencia> lstIncidencias = objRespuesta.lstIncidencias.Select(s =>
            new Incidencia
            {
                Columna = s.Columna,
                Linea = s.Linea,
                Mensaje = (s.Mensaje != null) ? (s.Mensaje.Length > 120) ? s.Mensaje.Substring(0, 120) : s.Mensaje : "",
                Valor = (s.Valor != null) ? (s.Valor.Length > 30) ? s.Valor.Substring(0, 30) : s.Valor : "",
                ValorEsperado = (s.ValorEsperado != null) ? (s.ValorEsperado.Length > 50) ? s.ValorEsperado.Substring(0, 50) : s.ValorEsperado : "",
            }).ToList();


            ConsoleTable.From(lstIncidencias)
                .Configure(o => o.NumberAlignment = Alignment.Right)
                .Write(Format.Alternative);

            Console.ReadKey();



            string strFileErrorName = Path.Combine(filePath, $"{strFileName} ERROR {DateTime.Now.ToString("dddd, dd MMMM yyyy hh mm ss tt")}.html");
            //File.WriteAllText(strFileErrorName, Util.ToHtmlString(lstMensajesFinales, objRespuesta.lstIncidencias), Encoding.UTF8);
        }

        private static bool CheckTargetDirectory(string checkDirectory)
        {
            bool blnContinuar = false;
            try
            {
                if (!Directory.Exists(checkDirectory))
                {
                    Directory.CreateDirectory(checkDirectory);
                }
                Console.WriteLine(checkDirectory);
                blnContinuar = true;
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return blnContinuar;
        }

        private static bool LecturaCatalogos(string strCatClaveProdServCPDir, string strCatClaveUnidadDir)
        {
            bool blnCorrecto = false;
            try
            {
                //Lectura de catalogo de Productos y servicios 
                using (StreamReader r = new StreamReader(strCatClaveProdServCPDir))
                {
                    string json = r.ReadToEnd();
                    lstClaveProdServCP = JsonConvert.DeserializeObject<List<ClaveProdServCP>>(json);
                }

                //Lectura de catalogo de Unidades
                using (StreamReader r = new StreamReader(strCatClaveUnidadDir))
                {
                    string json = r.ReadToEnd();
                    lstCatClaveUnidad = JsonConvert.DeserializeObject<List<ClaveUnidad>>(json);
                }

                if (lstClaveProdServCP != null && lstClaveProdServCP.Any() && lstCatClaveUnidad != null && lstCatClaveUnidad.Any())
                {
                    blnCorrecto = true;
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error.Message);
                throw;
            }
            return blnCorrecto;
        }

        private static void SetFullsize()
        {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3); //SW_MAXIMIZE = 3
        }

    }

}
