using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;

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

                        #region FechaEncabezado

                        // Texto con la fecha actual
                        var FechaActual = $"Fecha: {DateTime.Now.ToLongDateString()}";

                        // Obtener la celda E4
                        var E4_Fecha = datosHoja1.Descendants<Cell>().Where(c => c.CellReference.Value.Equals("E4")).FirstOrDefault();

                        E4_Fecha.RemoveAllChildren();
                        E4_Fecha.AppendChild(new InlineString(new Text { Text = FechaActual }));
                        E4_Fecha.DataType = CellValues.InlineString;

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
            }
        }

    }
}
