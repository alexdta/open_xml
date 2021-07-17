using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace open_xml
{
    public class office_excel
    {

        #region DatosEjemplo

        private static Random gen = new Random();

        private static DateTime RandomDay()
        {
            DateTime start = new DateTime(1995, 1, 1);
            int range = (DateTime.Today - start).Days;
            return start.AddDays(gen.Next(range));
        }

        private static DataTable datosEjemplo()
        {
            DataTable lEjemplo = new DataTable();
            lEjemplo.Columns.Add("Encabezado1");
            lEjemplo.Columns.Add("Encabezado2");
            lEjemplo.Columns.Add("Encabezado3");
            lEjemplo.Columns.Add("Encabezado4");
            lEjemplo.Columns.Add("Encabezado5");
            lEjemplo.Columns.Add("Encabezado6");
            lEjemplo.Columns.Add("Encabezado7");
            lEjemplo.Columns.Add("Encabezado8");
            lEjemplo.Columns.Add("Encabezado9");
            lEjemplo.Columns.Add("Encabezado10");
            lEjemplo.Columns.Add("Encabezado11");

            for (int i = 1; i <= 500; i++)
            {
                DataRow lFila = lEjemplo.NewRow();
                lFila["Encabezado1"] = $"Cedula-{i}";
                lFila["Encabezado2"] = $"Nombre1";
                lFila["Encabezado3"] = $"Apellido1";
                lFila["Encabezado4"] = $"Apellido2-{i}";
                lFila["Encabezado5"] = $"Dirección-{i}";
                lFila["Encabezado6"] = $"Telefono-{i}";
                lFila["Encabezado7"] = $"Celular-{i}";
                lFila["Encabezado8"] = $"Contacto-{i}";
                lFila["Encabezado9"] = $"Edad-{i}";
                lFila["Encabezado10"] = $"Estado Civil-{i}";
                lFila["Encabezado11"] = RandomDay().ToString("yyyy-MM-dd");

                lEjemplo.Rows.Add(lFila);
            }

            return lEjemplo;
        }

        #endregion

        /// <summary>
        /// Crea una celda en una posicion especificada con un valor definido
        /// </summary>
        /// <param name="posicion"></param>
        /// <param name="valor"></param>
        /// <returns></returns>
        private static Cell crearCelda(string posicion, string valor)
        {
            Cell celda = new Cell()
            {
                DataType = CellValues.String,
                CellReference = posicion,
                CellValue = new CellValue(valor)
            };

            return celda;
        }

        /// <summary>
        /// Crea una celda en una posicion especificada con un valor definido
        /// El texto se agregará a la SharedStringTable si no estuviese ya
        /// </summary>
        /// <param name="posicion"></param>
        /// <param name="valor"></param>
        /// <param name="Tabla"></param>
        /// <returns></returns>
        private static Cell crearCelda(string posicion, string valor, SharedStringTablePart Tabla)
        {
            var indice = AgregarSharedString(Tabla, valor);

            Cell celda = new Cell()
            {
                DataType = new EnumValue<CellValues>(CellValues.SharedString),
                CellReference = posicion,
                CellValue = new CellValue(indice)
            };

            return celda;
        }

        /// <summary>
        /// Verifica que el libro tenga una SharedStringTablePart
        /// Crea o retorna la existente según el caso
        /// </summary>
        /// <param name="Principal"></param>
        /// <returns></returns>
        private static SharedStringTablePart TablaStrings(WorkbookPart Principal)
        {
            if (Principal.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                return Principal.SharedStringTablePart;
            }

            return Principal.AddNewPart<SharedStringTablePart>();
        }

        /// <summary>
        /// Agrega el texto indicado a la SharedStringTable
        /// Si el texto ya existe retorna el indice
        /// </summary>
        /// <param name="Tabla"></param>
        /// <param name="Texto"></param>
        /// <returns></returns>
        private static int AgregarSharedString(SharedStringTablePart Tabla, string Texto)
        {
            if (Tabla.SharedStringTable == null)
            {
                Tabla.SharedStringTable = new SharedStringTable();
            }

            if (Tabla.SharedStringTable.Elements<SharedStringItem>().Where(e => e.InnerText.Equals(Texto)).Count() > 0)
            {
                return Tabla.SharedStringTable.Elements<SharedStringItem>()
                    .Select(e => e.InnerText)
                    .ToList()
                    .IndexOf(Texto);
            }

            Tabla.SharedStringTable.AppendChild(new SharedStringItem(new Text(Texto)));
            Tabla.SharedStringTable.Save();

            return Tabla.SharedStringTable.Count() - 1;
        }

        /// <summary>
        /// Crea una celda en una posicion especificada con un valor definido
        /// Se define si el campo va en negrita (bold)
        /// Opcionalmente se puede establecer el tamaño de la fuente
        /// </summary>
        /// <param name="posicion"></param>
        /// <param name="valor"></param>
        /// <param name="bold"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        private static Cell crearCelda(string posicion, string valor, bool bold, int fontSize = 12)
        {
            Cell celda = new Cell()
            {
                DataType = CellValues.InlineString,
                CellReference = posicion
            };

            RunProperties runProperties = new RunProperties();

            if (bold) { runProperties.Append(new Bold()); }
            runProperties.Append(new FontSize() { Val = fontSize });

            Run run = new Run();
            run.Append(new Text(valor));

            run.RunProperties = runProperties;

            InlineString inlineString = new InlineString();
            inlineString.Append(run);

            celda.Append(inlineString);

            return celda;
        }

        /// <summary>
        /// Cambia el texto de una celda especifica en una hoja
        /// </summary>
        /// <param name="datosHoja"></param>
        /// <param name="numeroCelda"></param>
        /// <param name="texto"></param>
        private static void TextoCelda(SheetData datosHoja, string numeroCelda, string texto)
        {
            //Buscar la celda en la hoja
            var celda =
                datosHoja.Descendants<Cell>()
                .Where(c => c.CellReference.Value.Equals(numeroCelda))
                .FirstOrDefault();

            //Si la celda ya contiene datos, estos se actualizan
            if (celda != null)
            {
                //Mejorar => Si el texto está en SharedStringTable
                celda.RemoveAllChildren();
                celda.AppendChild(new InlineString(new Text { Text = texto }));
                celda.DataType = CellValues.InlineString;
            }
            //Si la celda no existe, se crea la celda con sus datos
            else
            {
                //Nueva Celda
                celda = crearCelda(numeroCelda, texto);

                //Numero de Fila
                var NumFila = Convert.ToUInt32(Regex.Match(numeroCelda, @"\d+").Value);

                //Obtener la fila
                var Fila = ObtenerFila(datosHoja, NumFila);

                //Las celdas deben estar en orden
                Cell celdaReferencia = null;

                if (Fila.Elements<Cell>().Count() > 0)
                {
                    //Revisar antes de cual celda insertar la nueva
                    foreach (Cell cell in Fila.Elements<Cell>())
                    {
                        if (cell.CellReference.Value.Length == numeroCelda.Length)
                        {
                            if (string.Compare(cell.CellReference.Value, numeroCelda, true) > 0)
                            {
                                celdaReferencia = cell;
                                break;
                            }
                        }
                    }
                }

                //Insertar la celda antes de la de referencia (si existiese)
                Fila.InsertBefore(celda, celdaReferencia);

            }
        }

        /// <summary>
        /// Obtiene el texto de una celda
        /// Si el texto se encuentra en la SharedStringTable o si está directamente en la celda
        /// </summary>
        /// <param name="principal"></param>
        /// <param name="datosHoja"></param>
        /// <param name="referenciaCelda"></param>
        /// <returns></returns>
        private static string TextoCelda(WorkbookPart principal, SheetData datosHoja, string referenciaCelda)
        {
            string texto = string.Empty;

            var celda =
                datosHoja.Descendants<Cell>()
                .Where(c => c.CellReference.Value.Equals(referenciaCelda))
                .FirstOrDefault();

            //si es un sharedstring
            if (celda.DataType.Value == CellValues.SharedString)
            {
                int.TryParse(celda.InnerText, out int id);
                var dato =
                        principal
                        .SharedStringTablePart
                        .SharedStringTable
                        .Elements<SharedStringItem>()
                        .ElementAtOrDefault(id);
                if (dato != null)
                    texto = dato.InnerText;
            }
            else
            {
                texto = celda.InnerText;
            }

            return texto;
        }

        /// <summary>
        /// Obtiene una fila especifica de una hoja según el numero indicado
        /// Si la fila no existe la crea y la inserta en la hoja
        /// </summary>
        /// <param name="datosHoja"></param>
        /// <param name="numFila"></param>
        /// <returns></returns>
        private static Row ObtenerFila(SheetData datosHoja, uint numFila)
        {
            var Fila = datosHoja.Elements<Row>()
                .Where(r => r.RowIndex == numFila)
                .FirstOrDefault();

            if (Fila == null)
            {
                Fila = new Row() { RowIndex = numFila };

                var FilaRef = datosHoja.Elements<Row>()
                            .Where(r => r.RowIndex == numFila - 1)
                            .FirstOrDefault();

                datosHoja.InsertAfter(Fila, FilaRef);
            }

            return Fila;
        }

        public static void editarXLSX()
        {
            string lTemplate = Path.Combine(Environment.CurrentDirectory, "excelTemplate.xlsx");

            string lLibroNuevo = Path.Combine(Environment.CurrentDirectory, "reporteNuevo.xlsx");

            if (File.Exists(lTemplate))
            {
                File.Copy(lTemplate, lLibroNuevo, true);

                if (File.Exists(lLibroNuevo))
                {
                    using (var libro = SpreadsheetDocument.Open(lLibroNuevo, true))
                    {
                        var Principal = libro.WorkbookPart;

                        var hoja1 = Principal.WorksheetParts.ElementAt(0);

                        var datosHoja1 = hoja1.Worksheet.Elements<SheetData>().First();

                        //var tablaStrings = TablaStrings(Principal);

                        #region Encabezado

                        //Consecutivo
                        TextoCelda(datosHoja1, "K4", "01-132456-2021");

                        // Texto con la fecha actual
                        var FechaActual = $"Fecha: {DateTime.Now.ToLongDateString()}";
                        TextoCelda(datosHoja1, "E4", FechaActual);

                        // Insertar Texto en una Fila/Celda Nueva
                        TextoCelda(datosHoja1, "M7", "Edgar Chaves");

                        #endregion

                        #region DatosTabla

                        var datosTabla = datosEjemplo();
                        var cantidadDatos = datosTabla.Rows.Count;

                        // Encabezado de la tabla
                        var numFilaBase = 8;

                        foreach (DataRow dato in datosTabla.Rows)
                        {
                            var filaReferencia = datosHoja1.Elements<Row>().Where(r => r.RowIndex == numFilaBase).FirstOrDefault();

                            Row nuevaFila = new Row()
                            {
                                RowIndex = (uint)(numFilaBase + 1)
                            };

                            nuevaFila.Append(crearCelda($"B{(numFilaBase + 1)}", dato.ItemArray[0].ToString(), true));
                            nuevaFila.Append(crearCelda($"C{(numFilaBase + 1)}", dato.ItemArray[1].ToString()));
                            nuevaFila.Append(crearCelda($"D{(numFilaBase + 1)}", dato.ItemArray[2].ToString()));
                            nuevaFila.Append(crearCelda($"E{(numFilaBase + 1)}", dato.ItemArray[3].ToString()));
                            nuevaFila.Append(crearCelda($"F{(numFilaBase + 1)}", dato.ItemArray[4].ToString()));
                            nuevaFila.Append(crearCelda($"G{(numFilaBase + 1)}", dato.ItemArray[5].ToString()));
                            nuevaFila.Append(crearCelda($"H{(numFilaBase + 1)}", dato.ItemArray[6].ToString()));
                            nuevaFila.Append(crearCelda($"I{(numFilaBase + 1)}", dato.ItemArray[7].ToString()));
                            nuevaFila.Append(crearCelda($"J{(numFilaBase + 1)}", dato.ItemArray[8].ToString()));
                            nuevaFila.Append(crearCelda($"K{(numFilaBase + 1)}", dato.ItemArray[9].ToString()));
                            nuevaFila.Append(crearCelda($"L{(numFilaBase + 1)}", dato.ItemArray[10].ToString()));

                            // Se agrega la nueva fila despues de la fila de referencia
                            filaReferencia.InsertAfterSelf(nuevaFila);

                            numFilaBase++;
                        }

                        //Actualizar los datos con formato de tabla.
                        if (hoja1.TableDefinitionParts.Count() > 0)
                        {
                            var tabla = hoja1.TableDefinitionParts.First();
                            tabla.Table.Reference = $"B8:L{cantidadDatos + 8}";
                            tabla.Table.Save();
                        }

                        #endregion

                        libro.Close();

                        #region AbrirArchivo

                        try
                        {
                            Process.Start(lLibroNuevo);
                        }
                        catch (Exception lOpen)
                        {
                            throw new Exception($"No se puede abrir el archivo\n{lOpen.Message}");
                        }

                        #endregion

                    }
                }
                else
                {
                    throw new Exception("No se pudo generar el nuevo archivo");
                }
            }
            else
            {
                throw new Exception("No se encuentra la plantilla");
            }
        }

    }
}
