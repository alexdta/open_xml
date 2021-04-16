using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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
                            var marcadoresDoc = cuerpoDoc.Descendants<BookmarkStart>();

                            var fecha = DateTime.Now;

                            TextoMarcador(marcadoresDoc, "dia", fecha.Day.ToString());
                            TextoMarcador(marcadoresDoc, "mes", fecha.ToString("MMMM", CultureInfo.CurrentCulture));
                            TextoMarcador(marcadoresDoc, "anho", fecha.Year.ToString());

                            TextoMarcador(marcadoresDoc, "nombre", "Edgar CM.");
                            TextoMarcador(marcadoresDoc, "notas", "Otros datos dentro de una tabla.");

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

                            #region Imagen Marcador

                            var marcadorImagen = marcadoresDoc.FirstOrDefault(bms => bms.Name == "imagen");
                            var imagen = Path.Combine(Environment.CurrentDirectory, "crash.jpg");

                            ImagenMarcador(doc, marcadorImagen, imagen);

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

        public static void ImagenMarcador(WordprocessingDocument pDocumento, BookmarkStart pMarcador, string pImagen)
        {
            // Eliminar todo dentro del marcador
            OpenXmlElement elem = pMarcador.NextSibling();
            while (elem != null && !(elem is BookmarkEnd))
            {
                OpenXmlElement nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }

            var imagePart = AgregarImagePart(pDocumento.MainDocumentPart, pImagen);

            #region Dimensiones Imagen

            // Calcular el tamaño según las dimensiones del archivo
            var img = new BitmapImage(new Uri(pImagen, UriKind.RelativeOrAbsolute));
            var anchoPx = img.PixelWidth;
            var altoPx = img.PixelHeight;
            var dpiHorizontal = img.DpiX;
            var dpiVertical = img.DpiY;
            const int emusPerInch = 914400;

            var anchoEmus = (long)(anchoPx / dpiHorizontal * emusPerInch);
            var altoEmus = (long)(altoPx / dpiVertical * emusPerInch);

            #endregion

            // Insertar imagen
            AgregarImagen(pDocumento.MainDocumentPart.GetIdOfPart(imagePart), pMarcador, anchoEmus, altoEmus);
        }

        public static ImagePart AgregarImagePart(MainDocumentPart mainPart, string imageFilename)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(imageFilename, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private static void AgregarImagen(string pIdPosicion, BookmarkStart pMarcador, long CX, long CY)
        {
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent()
                        {
                            Cx = CX,
                            Cy = CY
                        },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = 1U, Name = "Imagen" },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = 0U,
                                            Name = "Imagen.jpg"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(new A.BlipExtensionList(new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                        {
                                            Embed = pIdPosicion,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                            new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = CX, Cy = CY }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                            {
                                Uri =
                                        "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            }))
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U,
                        EditId = "50D07946"
                    });

            pMarcador.Parent.InsertAfter(new Run(element), pMarcador);
        }

    }
}