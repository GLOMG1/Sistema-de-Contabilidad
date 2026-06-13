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
                tbl = sheet.ListObjects["tblEgresos"];
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
        public List<string> Encabezados(string Nombretabla)
        {
            Excel.Worksheet hoja;
            Excel.ListObject tbl;
            List<string> campos = new List<string>();

            switch (Nombretabla)
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
                    return new List<string>();
            }
            tbl = hoja.ListObjects[Nombretabla];
            Excel.Range headerRange = (Excel.Range)tbl.HeaderRowRange;
            int totalColumnas = headerRange.Columns.Count;


            for (int i = 1; i <= totalColumnas; i++)
            {
                Excel.Range celda = (Excel.Range)headerRange.Cells[1, i];
                campos.Add(celda.Value2?.ToString() ?? ""); 
            }

            return campos;
        }
        public List<ClassBancos> ObtenerBancos(bool tipoMov, List<Filtro>PanelSuperior)
        {
            List<ClassBancos> bancos = new List<ClassBancos>();

            Excel.Worksheet sheet = HojaActiva();
            Excel.ListObject tbl;
            try
            {
                tbl = sheet.ListObjects["EstadoCuenta"];
            }
            catch
            {
                MessageBox.Show("No se encontró la tabla Bancos.");
                return bancos;
            }
            object[,] datos = (object[,])tbl.DataBodyRange.Value2;

            object[,] encabezados = (object[,])tbl.HeaderRowRange.Value2;
            Dictionary<string, int> columnas = new Dictionary<string, int>();

            for (int i = 1; i <= encabezados.GetLength(1); i++)
            {
                columnas[encabezados[1, i]?.ToString() ?? ""] = i;
            }

            int filas = datos.GetLength(0);

            for (int fila = 1; fila <= filas; fila++)
            {
                if (PanelSuperior != null && PanelSuperior.Any())
                {
                    bool validador = true;
                    foreach (Filtro x in PanelSuperior)
                    {
                        if (!columnas.ContainsKey(x.Campo)) continue;

                        int idx = columnas[x.Campo];
                        string valorCelda = datos[fila, idx]?.ToString() ?? "";

                        if (!valorCelda.ToLower().Contains(x.Valor.ToLower()))
                        {
                            validador = false;
                            break;
                        }
                    }
                    if (!validador) continue;
                }
                if (tipoMov) {
                    if (Convert.ToDecimal(datos[fila, 6]) == 0)
                        continue;
                }
                else {
                    if (Convert.ToDecimal(datos[fila, 7]) == 0)
                        continue; 
                }
                if (datos[fila, 19]?.ToString() == "CONCILIADO")
                    continue;

                bancos.Add(new ClassBancos
                {
                    IdDataBody = fila,
                    IdRegistro = Convert.ToInt32(datos[fila, 1] ?? 0),
                    FechaOperacion = DateTime.FromOADate(Convert.ToDouble(datos[fila, 2])),
                    FechaLiquidacion = datos[fila, 3]?.ToString(),
                    Concepto = datos[fila, 4]?.ToString(),
                    Referencia = datos[fila, 5]?.ToString(),
                    Importe = tipoMov 
                            ? Convert.ToDecimal(datos[fila, 6] ?? 0):
                            Convert.ToDecimal(datos[fila, 7] ?? 0),
                    SaldoOperacion = Convert.ToDecimal(datos[fila, 8] ?? 0),
                    SaldoLiquidacion = Convert.ToDecimal(datos[fila, 9] ?? 0),
                    Año = Convert.ToInt32(datos[fila, 10] ?? 0),
                    Mes = datos[fila, 11]?.ToString(),
                    Banco = datos[fila, 12]?.ToString(),
                    NumCuenta = datos[fila, 13]?.ToString(),
                    Clasificacion = datos[fila, 14]?.ToString(),
                    SerieFolioInternoCFDI = datos[fila, 15]?.ToString(),
                    FechaCFDI = datos[fila, 16]?.ToString(),
                    FolioUUIDCFDI = datos[fila, 17]?.ToString(),
                    NombreEmisorReceptorCFDI = datos[fila, 18]?.ToString(),
                    Estatus = datos[fila, 19]?.ToString(),
                    IdRegistroAuxiliar = datos[fila, 20]?.ToString(),
                    FechaAuxiliar = datos[fila, 21]?.ToString(),
                    FolioAuxilar = datos[fila, 22]?.ToString(),
                    EstatusAUXILIAR = datos[fila, 23]?.ToString()
                });
            }
            return bancos;
        }
        public List<ClassFacturas> ObtenerFacturas(string Nombretabla)
        {
            List<ClassFacturas> tblFacturas = new List<ClassFacturas>();
            Excel.Worksheet hoja;
            Excel.ListObject tbl;
            switch (Nombretabla)
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
                    return tblFacturas;
            }
            tbl = hoja.ListObjects[Nombretabla];
            object[,] datos = (object[,])tbl.DataBodyRange.Value2;

            int numFilas = datos.GetLength(0);

            for(int fila = 1; fila <= numFilas; fila++)
            {
                tblFacturas.Add(new ClassFacturas
                {
                    FechaCFDI = DateTime.FromOADate(Convert.ToDouble(datos[fila, 12])),
                    FolioUUID = datos[fila, 16]?.ToString(),
                    SerieFolioInterno = datos[fila, 7]?.ToString(),
                    RfcEmisor = datos[fila, 8]?.ToString(),
                    NombreEmisor = datos[fila, 9]?.ToString(),
                    TipoCFDI = datos[fila, 13]?.ToString(),
                    FormaPago = datos[fila, 21]?.ToString(),
                    Concepto = datos[fila, 24]?.ToString(),
                    Fiscal = datos[fila, 5]?.ToString(),
                    Estatus = datos[fila, 71]?.ToString(),
                });
            }
            return tblFacturas;
        }
    }
}
