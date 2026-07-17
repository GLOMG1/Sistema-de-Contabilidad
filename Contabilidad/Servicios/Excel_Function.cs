using Contabilidad.Formularios;
using Contabilidad.Modelos;
using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Contabilidad.Servicios
{
    public class tblsExcel
    {
        public string nombre { get; set; }
        public string celdaInicio { get; set; }
        public List<columnas> columnas { get; set; }
    }
    public class columnas
    {
        public string nombre { get; set; }
        public string type { get; set; }
        public int width { get; set; }
        public string formato { get; set; }
    }
    internal class Excel_Function
    {
        private Excel.Application app;

        public Excel_Function()
        {
            app = Globals.ThisAddIn.Application;
        }
        public Excel.Worksheet HojaActiva()
        {
            return (Excel.Worksheet)app.ActiveSheet;
        }
        public void CrearTabla(Contribuyentes objContribuyente)
        {
            Excel.Worksheet sheet = HojaActiva();

            bool papelTrabajo = sheet.ListObjects
                .Cast<Excel.ListObject>()
                .Any(t => t.Name == "tblEgresos");

            if (!papelTrabajo)
            {
                FuncionesExcel(true);

                Excel.Range titulo = (Excel.Range)sheet.Cells[2, 2];
                titulo.Value = objContribuyente.Nombre;
                titulo.Font.Size = 14;
                titulo.Font.Name = "Arial";
                titulo.Font.Bold = true;

                titulo = (Excel.Range)sheet.Cells[3, 2];
                titulo.Value = objContribuyente.Rfc;
                titulo.Font.Size = 10;
                titulo.Font.Name = "Arial";

                titulo = (Excel.Range)sheet.Cells[4, 2];
                titulo.Value = objContribuyente.Regimen;
                titulo.Font.Size = 10;
                titulo.Font.Name = "Arial";
                titulo.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                string json = File.ReadAllText("PapelesDeTrabajo.json");

                tblsExcel templante = JsonSerializer.Deserialize<tblsExcel>(json);

                for (int i = 0; i < templante.columnas.Count; i++)
                {
                    var col = templante.columnas[i];
                    Excel.Range cell = (Excel.Range)sheet.Cells[6, i + 1];

                    cell.Value = col.nombre;
                    cell.ColumnWidth = col.width;

                    if (!string.IsNullOrEmpty(col.formato))
                        cell.NumberFormat = col.formato;
                }

                Excel.Range rango = sheet.Range["B6:AO7"];

                Excel.ListObject tabla = sheet.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    rango,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes
                    );
                tabla.Name = templante.nombre;
                tabla.TableStyle = "TableStyleLight21";
            }
            FuncionesExcel(false);
        }
        public void CargarDatos()
        {
            Excel.Worksheet sheet = HojaActiva();
            Excel.ListObject tbl;
            try
            {
                tbl = sheet.ListObjects["TblFactEgresos"];
            }
            catch
            {
                return;
            }

            gestor_Archivos CarpetaXML = new gestor_Archivos();
            IEnumerable<string> ConjuntoXML = CarpetaXML.ArchivosXML();

            if (ConjuntoXML == null)
                return;
            FuncionesExcel(true);
            object[,] ColumnaUUID = (object[,])tbl.ListColumns["Folio UUID CFDI"].DataBodyRange.Value2;

            int contador = 1;

            foreach (string XML in ConjuntoXML)
            {
                app.StatusBar = $"Archivos procesados: {contador} de {ConjuntoXML.Count()}";

                xml_Function Archivo = new xml_Function();

                var CFDI = Archivo.leerArchivos(XML);
                if (CFDI == null)
                    continue;

                Herramientas Rango = new Herramientas();
                object[] DatosFactura = Rango.ContructorFila(CFDI, tbl);

                if (FiltroRepeticiones(ColumnaUUID, DatosFactura[Rango.IdFolioUUID]))
                {
                    switch (DatosFactura[Rango.TipoFact])
                    {
                        case "I":
                            Excel.ListRow nuevaFila = tbl.ListRows.Add();
                            nuevaFila.Range.Value2 = DatosFactura;
                            break;
                        case "P":
                            foreach (object x in Rango.Añadir)
                            {
                                Excel.ListRow P = tbl.ListRows.Add();
                                P.Range.Value2 = x;
                            }
                            break;
                    }   
                }

                contador++;
            }
            FuncionesExcel(false);
        }
        public void EscribirDatos(string nombreTabla, int numFila, string columna, string valor)
        {
            Excel.Worksheet hoja;
            switch (nombreTabla)
            {
                case "EstadoCuenta":
                hoja = (Excel.Worksheet)app.Worksheets["Bancos"];
                break;
                case "TblFactEgresos":
                    hoja = (Excel.Worksheet)app.Worksheets["BaseEgresos"];
                    break;
                case "TblFactIngresos":
                    hoja = (Excel.Worksheet)app.Worksheets["BaseIngresos"];
                    break;
                default:
                    throw new ArgumentException($"Tabla '{nombreTabla}' no reconocida.");
            }

            var tbl = hoja.ListObjects[nombreTabla];
            var columnaExcel = tbl.ListColumns[columna];
            columnaExcel.DataBodyRange[numFila] = valor;
        }
        private void FuncionesExcel(bool interruptor)
        {

            if (interruptor)
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;
            }
            else
            {
                app.ScreenUpdating = true;
                app.EnableEvents = true;
                app.DisplayAlerts = true;
                app.StatusBar = false;
            }
        }
        private bool FiltroRepeticiones(object[,] datos, object fact)
        {
            if (datos == null)
                return true;

            string StrFact = fact?.ToString();
            for(int i = 1; i <= datos.GetLength(0); i++)
            {
                if (datos[i,1]?.ToString() == StrFact)
                    return false;
            }
            return true;
        }
        public void FuncionDePrueba()
        {
            string json = gestor_Archivos.AbrirInfoClientes();
            RutaDeTrabajo Prueba = JsonSerializer.Deserialize<RutaDeTrabajo>(json);
            MessageBox.Show(Prueba.rutaPrincipal);
        }
        public ExcelDatosTabla ObtenerTabla(string nombreTabla)
        {
            Excel.Worksheet hoja;

            switch (nombreTabla)
            {
                case "EstadoCuenta":
                    hoja = (Excel.Worksheet)app.Worksheets["Bancos"];
                    break;
                case "TblFactEgresos":
                    hoja = (Excel.Worksheet)app.Worksheets["BaseEgresos"];
                    break;
                case "TblFactIngresos":
                    hoja = (Excel.Worksheet)app.Worksheets["BaseIngresos"];
                    break;
                default:
                    throw new ArgumentException($"Tabla '{nombreTabla}' no reconocida.");
            }

            var tbl = hoja.ListObjects[nombreTabla];

            var headersRaw = (object[,])tbl.HeaderRowRange.Value2;
            var headers = new List<string>();
            for (int i = 1; i <= headersRaw.GetLength(1); i++)
                headers.Add(headersRaw[1, i]?.ToString());


            return new ExcelDatosTabla
            {
                NombreTabla = nombreTabla,
                Encabezados = headers,
                Datos = (object[,])tbl.DataBodyRange.Value2
            };
        }
    }
}
