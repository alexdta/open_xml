using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

namespace open_xml
{
    public class office_word
    {
        private static void TextoMarcador(IEnumerable<BookmarkStart> pListMarcadores, string pMarcador, string pTexto)
        {
            var marcador = pListMarcadores.FirstOrDefault(bms => bms.Name == pMarcador);
            if (marcador != null)
            {
                var marcadorRun = marcador.NextSibling<Run>();
                if (marcadorRun != null)
                {
                    marcadorRun.GetFirstChild<Text>().Text = pTexto;
                }
            }
        }

        public static void TextoCelda(TableCell celda, string dato)
        {
            var Run = new Run();
            var PropiedadesRun = new RunProperties();
            PropiedadesRun.RunFonts = new RunFonts() { Ascii = "Century Gothic" };
            PropiedadesRun.FontSize = new FontSize() { Val = "20" }; //Doble del deseado 20 => 10
            Run.Append(PropiedadesRun);
            Run.Append(new Text(dato));

            var PropiedadesParra = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
            var Parrafo = new Paragraph(PropiedadesParra);
            Parrafo.Append(Run);

            celda.RemoveAllChildren<Paragraph>();
            celda.Append(Parrafo);
        }

        public static void editarDoc()
        {
            try
            {
                string lTemplate = Path.Combine(Environment.CurrentDirectory, "openXmlTemplate.docx");

                string lNuevoDoc = Path.Combine(Environment.CurrentDirectory, "nuevoDoc.docx");

                if (File.Exists(lTemplate))
                {
                    File.Copy(lTemplate, lNuevoDoc, true);

                    if (File.Exists(lNuevoDoc))
                    {
                        using (var doc = WordprocessingDocument.Open(lNuevoDoc, true))
                        {
                            MainDocumentPart Principal = doc.MainDocumentPart;

                            #region Encabezado

                            var encabezado = Principal.HeaderParts.First();
                            var marcadoresEncabezado = encabezado.RootElement.Descendants<BookmarkStart>();

                            TextoMarcador(marcadoresEncabezado, "consecutivo", "01-0789-2021");

                            #endregion

                            #region Cuerpo_Documento

                            var cuerpoDoc = Principal.RootElement;
                            var marcadoresDoc2 = cuerpoDoc.Descendants<BookmarkStart>();

                            var fecha = DateTime.Now;

                            TextoMarcador(marcadoresDoc2, "dia", fecha.Day.ToString());
                            TextoMarcador(marcadoresDoc2, "mes", fecha.ToString("MMMM", CultureInfo.CurrentCulture));
                            TextoMarcador(marcadoresDoc2, "anho", fecha.Year.ToString());

                            TextoMarcador(marcadoresDoc2, "nombre", "Edgar CM.");
                            TextoMarcador(marcadoresDoc2, "notas", "Otros datos dentro de una tabla.");

                            #endregion

                            #region Tabla

                            var tabla = cuerpoDoc.Descendants<Table>().ElementAt(1);

                            // Fila base
                            var ultimaFila = tabla.Elements<TableRow>().Last();

                            for (int j = 0; j < 500; j++)
                            {
                                TableRow nuevaFila = (TableRow)ultimaFila.CloneNode(true);
                                var Celdas = nuevaFila.Descendants<TableCell>();

                                var Elemento = new ElementoTabla(j);

                                TextoCelda(Celdas.ElementAt(0), Elemento.Codigo);
                                TextoCelda(Celdas.ElementAt(1), Elemento.Cantidad.ToString());
                                TextoCelda(Celdas.ElementAt(2), Elemento.Costo.ToString());
                                TextoCelda(Celdas.ElementAt(3), Elemento.Total.ToString());

                                tabla.AppendChild(nuevaFila);
                            }

                            tabla.RemoveChild(ultimaFila);

                            #endregion

                            #region Eliminar Tabla

                            var tablaEliminar = cuerpoDoc.Descendants<Table>().ElementAt(2);
                            tablaEliminar.Remove();

                            #endregion

                            #region Eliminar Fila

                            var tablaEliminarFila = cuerpoDoc.Descendants<Table>().ElementAt(2); //Al eliminar la tabla anterior, el indice cambia
                            var filaEliminar = tablaEliminarFila.Descendants<TableRow>().ElementAt(2);
                            filaEliminar.Remove();

                            #endregion

                            doc.Close();
                        }

                        try
                        {
                            Process.Start(lNuevoDoc);
                        }
                        catch (Exception lOpen)
                        {
                            throw new Exception($"No se puede abrir el archivo\n{lOpen.Message}");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}