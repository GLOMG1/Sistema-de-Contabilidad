using Contabilidad.Modelos;
using Contabilidad.Servicios;
using Microsoft.SqlServer.Server;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;


namespace Contabilidad.Formularios
{
    public class Filtro
    {
        public string Campo { get; set;  }
        public string Valor {  get; set; }
    }

    public partial class VentanaConciliacion : Window
    {
        /// Elementos de uso dentro de la venta
        Excel_Function excel = new Excel_Function();
        public ExcelDatosTabla tablaBancos { get; set; }
        public ExcelDatosTabla tablaEgresos { get; set;  }
        public ExcelDatosTabla tablaIngresos { get; set; }
        public List<Filtro> FiltrosBancos { get; set; }
        public List<Filtro> FiltrosFacturas {  get; set; }
        
        private DispatcherTimer timerFacturas;
        private DispatcherTimer timerBancos;

        /// Metodos esenciales dentro de la ventana
        public VentanaConciliacion()
        {
            InitializeComponent();

            IniciarTimers();

            tablaBancos = excel.ObtenerTabla("EstadoCuenta");
            tablaEgresos = excel.ObtenerTabla("TblFactEgresos");
            tablaIngresos = excel.ObtenerTabla("TblFactIngresos");

            InicializarComboBox();

            FiltrosBancos = new List<Filtro>();
            FiltrosFacturas = new List<Filtro>();
        }
        public void MostarDataGrid_Cheked(object sender, RoutedEventArgs e)
        {
            LimpiarBusquedas();
            FiltrosBancos.Clear();
            FiltrosFacturas.Clear();

            if (rb_E.IsChecked == true)
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, false);
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaEgresos);
            }
            if (rb_I.IsChecked == true)
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, true);
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaIngresos);
            }
        }
        public void Btn_Buscar(object sender, RoutedEventArgs e)
        {
            if (dgBancos.SelectedItems.Count == 0)
            {
                MessageBox.Show("Olvidaste Seleccionar un movimiento bancario");
                return;
            }

            decimal ImporteBuscado = 0;
            foreach (ClassBancos item in dgBancos.SelectedItems)
            {
                ImporteBuscado = ImporteBuscado + item.Importe;
            }

            dgFacturas.SelectedItems.Clear();
            foreach (ClassFacturas item in dgFacturas.Items)
            {
                if(item.TotalCFDI == ImporteBuscado)
                {
                    dgFacturas.SelectedItems.Add(item);
                    dgFacturas.ScrollIntoView(item);
                    return;
                }
            }
        }
        public void Btn_Conciliar(object sender, RoutedEventArgs e)
        {
            if (dgBancos.SelectedItems.Count == 0)
            {
                MessageBox.Show("Olvidaste Seleccionar un movimiento bancario");
                return;
            }
            if (dgFacturas.SelectedItems.Count == 0)
            {
                MessageBox.Show("Olvidaste Seleccionar alguna factura");
                return;
            }

            string nombreTabla = rb_E.IsChecked == true ? "TblFactEgresos": "TblFactIngresos";

            List<ClassBancos> filasBancos = dgBancos.SelectedItems.Cast<ClassBancos>().ToList();
            List<ClassFacturas> filasFacturas = dgFacturas.SelectedItems.Cast<ClassFacturas>().ToList();

            /// tabla Facturas
            string IdRegistroBancos = string.Join("; ", filasBancos.Select(f => f.IdDataBody).Distinct()) + "; ";
            string FechaBancos = string.Join("; ", filasBancos.Select(f => f.FechaOperacion.ToString("dd/MM/yyyy")).Distinct()) + "; ";
            string FolioBancos; 
            string Estatus = "CONCILIADO";
            string Banco = string.Join("; ", filasBancos.Select(f => f.Banco).Distinct());
            string NumeroCuenta = string.Join("; ", filasBancos.Select(f => f.NumCuenta).Distinct());

            /// tabla Bancos

            string SerieFolioInternoCFDI = string.Join("; ", filasFacturas.Select(b => b.SerieFolioInterno).Distinct()) + "; ";
            string FechaCFDI = string.Join("; ", filasFacturas.Select(b => b.FechaCFDI.ToString("dd/MM/yyyy")).Distinct()) + "; ";
            string FolioUUIDCFDI = string.Join("; ", filasFacturas.Select(b => b.FolioUUID).Distinct()) + "; ";
            string NombreEmisorReceptorCFDI = string.Join("; ", filasFacturas.Select(b => b.NombreEmisor).Distinct()) + "; ";

            /// aplicar formato a las facturas PPD:
            filasFacturas = busquedaComplementos(filasFacturas, nombreTabla);

            foreach (ClassFacturas f in filasFacturas)
            {
                excel.EscribirDatos(nombreTabla, f.filaTabla, "Id Registro Bancos", IdRegistroBancos);
                excel.EscribirDatos(nombreTabla, f.filaTabla, "Fecha Bancos", FechaBancos);
                excel.EscribirDatos(nombreTabla, f.filaTabla, "Estatus", Estatus);
                excel.EscribirDatos(nombreTabla, f.filaTabla, "Banco", Banco);
                excel.EscribirDatos(nombreTabla, f.filaTabla, "Num Cuenta", NumeroCuenta);
            }

            foreach(ClassBancos b in filasBancos)
            {
                excel.EscribirDatos("EstadoCuenta", b.filaTabla, "Serie Folio Interno CFDI", SerieFolioInternoCFDI);
                excel.EscribirDatos("EstadoCuenta", b.filaTabla, "Fecha CFDI", FechaCFDI);
                excel.EscribirDatos("EstadoCuenta", b.filaTabla, "Folio UUID CFDI", FolioUUIDCFDI);
                excel.EscribirDatos("EstadoCuenta", b.filaTabla, "Nombre Emisor Receptor CFDI", NombreEmisorReceptorCFDI);
                excel.EscribirDatos("EstadoCuenta", b.filaTabla, "Estatus", Estatus);
            }

            tablaBancos = excel.ObtenerTabla("EstadoCuenta");
            tablaEgresos = excel.ObtenerTabla("TblFactEgresos");
            tablaIngresos = excel.ObtenerTabla("TblFactIngresos");

            if (rb_E.IsChecked == true)
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, false);
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaEgresos);
            }
            if (rb_I.IsChecked == true)
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, true);
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaIngresos);
            }

            labelInferior.Content = "0.00";
            labelSuperior.Content = "0.00";
        }
        public void Btn_GenerarCheque(object sender, RoutedEventArgs e)
        {
            if (rb_E.IsChecked == false)
                return;

            List<ClassFacturas> filasFacturas = dgFacturas.SelectedItems.Cast<ClassFacturas>().ToList();

            foreach(ClassFacturas f in filasFacturas)
            {
                excel.EscribirDatos("TblFactEgresos", f.filaTabla, "Folio Bancos", "CH-0000000");
                excel.EscribirDatos("TblFactEgresos", f.filaTabla, "Estatus", "CONCILIADO");
            }

            tablaEgresos = excel.ObtenerTabla("TblFactEgresos");
            dgFacturas.ItemsSource = CrearDataGridFacturas(tablaEgresos);
            labelInferior.Content = "0.00";
        }
        public void SumarSeleccion_DG(object sender, SelectionChangedEventArgs e)
        {
            decimal contador = 0;
            foreach(ClassBancos item in dgBancos.SelectedItems)
            {
                contador = contador + item.Importe;
                labelSuperior.Content = contador.ToString("$#,##0.00");
            }

            contador = 0;
            foreach (ClassFacturas item in dgFacturas.SelectedItems)
            {
                contador = contador + item.TotalCFDI;
                labelInferior.Content = contador.ToString("$#,##0.00");
            }
        }
        public void CambiosValoresBancos(object sender, TextChangedEventArgs e)
        {
            timerBancos.Stop(); 
           timerBancos.Start(); 
        }
        public void CambiosValoresFacturas(object sender, TextChangedEventArgs e)
        {
            timerFacturas.Stop();
            timerFacturas.Start();
        }
        /// Funciones para Metodos
        private void InicializarComboBox()
        {
            Superior_Campo1.ItemsSource = tablaBancos.Encabezados;
            Superior_Campo2.ItemsSource = tablaBancos.Encabezados;
            Superior_Campo3.ItemsSource = tablaBancos.Encabezados;
            Superior_Campo4.ItemsSource = tablaBancos.Encabezados;

            Inferior_Campo1.ItemsSource = tablaEgresos.Encabezados;
            Inferior_Campo2.ItemsSource = tablaEgresos.Encabezados;
            Inferior_Campo3.ItemsSource = tablaEgresos.Encabezados;
            Inferior_Campo4.ItemsSource = tablaEgresos.Encabezados;
        }
        private void IniciarTimers()
        {
            timerFacturas = new DispatcherTimer();
            timerFacturas.Interval = TimeSpan.FromMilliseconds(650);
            timerFacturas.Tick += busquedaFiltrosFactura;

            timerBancos = new DispatcherTimer();
            timerBancos.Interval = TimeSpan.FromMilliseconds(650);
            timerBancos.Tick += busquedaFiltrosBanco;
        }
        private void LimpiarBusquedas()
        {
            var controles = new Control[]
            {
                Superior_Campo1, Superior_Campo2, Superior_Campo3, Superior_Campo4,
                Superior_Valor1, Superior_Valor2, Superior_Valor3, Superior_Valor4,
                Inferior_Campo1, Inferior_Campo2, Inferior_Campo3, Inferior_Campo4,
                Inferior_Valor1, Inferior_Valor2, Inferior_Valor3, Inferior_Valor4
            };

            foreach (var control in controles)
            {
                switch (control)
                {
                    case TextBox tb:
                        tb.Clear();
                        break;

                    case ComboBox cb:
                        cb.SelectedIndex = -1;
                        break;
                }
            }
        }
        private List<ClassFacturas> CrearDataGridFacturas(ExcelDatosTabla tabla)
        {
            List<ClassFacturas> tblFiltrada = new List<ClassFacturas>();
            for (int fila = 1; fila <= tabla.NumFilas; fila++)
            {
                bool cumpleFiltros = true;

                foreach (Filtro filtro in FiltrosFacturas)
                {
                    string valorTabla = tabla.GetValor(fila, filtro.Campo)?.ToString() ?? "";

                    if (valorTabla.IndexOf(filtro.Valor, StringComparison.OrdinalIgnoreCase) < 0)
                        cumpleFiltros = false;
                }
                if (tabla.GetValor(fila, "Forma Pago CFDI")?.ToString() == "99" ||
                    tabla.GetValor(fila, "Estatus")?.ToString() == "CONCILIADO" ||
                    tabla.GetValor(fila, "Tipo CFDI")?.ToString() == "N" ||
                    tabla.GetValor(fila, "Fiscal")?.ToString() == "CANCELADA")
                    cumpleFiltros = false;


                if (!cumpleFiltros)
                    continue;

                tblFiltrada.Add(new ClassFacturas
                {
                    filaTabla = fila,
                    FechaCFDI = DateTime.FromOADate(Convert.ToDouble(tabla.GetValor(fila, "Fecha CFDI"))),
                    FolioUUID = tabla.GetValor(fila, "Folio UUID CFDI")?.ToString(),
                    RelacionNC = tabla.GetValor(fila, "Folio UUID CFDI Relacionado NC")?.ToString(),
                    RelacionREP = tabla.GetValor(fila, "Folio UUID CFDI Relacionado REP")?.ToString(),
                    SerieFolioInterno = tabla.GetValor(fila, "Serie Folio Interno CFDI")?.ToString(),
                    RfcEmisor = tabla.GetValor(fila, "RFC Emisor CFDI")?.ToString(),
                    NombreEmisor = tabla.NombreTabla == "TblFactEgresos" ? tabla.GetValor(fila, "Nombre Emisor CFDI")?.ToString() : tabla.GetValor(fila, "Nombre Receptor CFDI")?.ToString(),
                    TipoCFDI = tabla.GetValor(fila, "Tipo CFDI")?.ToString(),
                    FormaPago = tabla.GetValor(fila, "Forma Pago CFDI")?.ToString(),
                    MetodoPago = tabla.GetValor(fila, "Metodo Pago CFDI")?.ToString(),
                    Concepto = tabla.GetValor(fila, "Concepto CFDI")?.ToString(),
                    Fiscal = tabla.GetValor(fila, "Fiscal")?.ToString(),
                    Estatus = tabla.GetValor(fila, "Estatus")?.ToString(),
                    TotalCFDI = Convert.ToDecimal(tabla.GetValor(fila, "TOTAL CFDI") ?? 0)
                });
            }
            return tblFiltrada;
        }
        private List<ClassBancos> CrearDataGridBancos(ExcelDatosTabla tabla, bool tipo)
        {
            List<ClassBancos> tblFiltrada = new List<ClassBancos>();
            for (int fila = 1; fila <= tabla.NumFilas; fila++)
            {
                bool cumpleFiltros = true;

                foreach (Filtro filtro in FiltrosBancos)
                {
                    string valorTabla = tabla.GetValor(fila, filtro.Campo)?.ToString() ?? "";

                    if (valorTabla.IndexOf(filtro.Valor, StringComparison.OrdinalIgnoreCase) < 0)
                        cumpleFiltros = false;
                }
                if (tipo)
                {
                    if (Convert.ToDecimal(tabla.GetValor(fila, "Depositos") ?? 0) == 0 ||
                        tabla.GetValor(fila, "Estatus")?.ToString() != null)
                        cumpleFiltros = false;
                }
                else
                {
                    if (Convert.ToDecimal(tabla.GetValor(fila, "Retiros") ?? 0) == 0 ||
                        tabla.GetValor(fila, "Estatus")?.ToString() != null)
                        cumpleFiltros = false;
                }

                if (!cumpleFiltros)
                    continue;

                tblFiltrada.Add(new ClassBancos
                {
                    filaTabla = fila,
                    IdDataBody =  Convert.ToInt32(tabla.GetValor(fila, "Id Registro")),
                    FechaOperacion = DateTime.FromOADate(Convert.ToDouble(tabla.GetValor(fila, "Fecha Operacion"))),
                    Concepto = tabla.GetValor(fila, "Concepto")?.ToString(),
                    Referencia = tabla.GetValor(fila, "Referencia")?.ToString(),
                    Importe = tipo ? Convert.ToDecimal(tabla.GetValor(fila, "Depositos") ?? 0) : Convert.ToDecimal(tabla.GetValor(fila, "Retiros") ?? 0),
                    Banco = tabla.GetValor(fila, "Banco")?.ToString(),
                    NumCuenta = tabla.GetValor(fila, "Num Cuenta")?.ToString(),
                    Clasificacion = tabla.GetValor(fila, "Clasificacion")?.ToString(),
                });
            }
            return tblFiltrada;
        }
        private void busquedaFiltrosFactura(object sender, EventArgs e)
        {
            if (FiltrosFacturas == null)
                return;

            FiltrosFacturas.Clear();

            /// Filtros predeterminados

            if (Inferior_Campo1.Text != "" && Inferior_Valor1.Text != "")
            {
                FiltrosFacturas.Add(new Filtro
                {
                    Campo = Inferior_Campo1.Text,
                    Valor = Inferior_Valor1.Text

                });
            }
            if (Inferior_Campo2.Text != "" && Inferior_Valor2.Text != "")
            {
                FiltrosFacturas.Add(new Filtro
                {
                    Campo = Inferior_Campo2.Text,
                    Valor = Inferior_Valor2.Text

                });
            }
            if (Inferior_Campo3.Text != "" && Inferior_Valor3.Text != "")
            {
                FiltrosFacturas.Add(new Filtro
                {
                    Campo = Inferior_Campo3.Text,
                    Valor = Inferior_Valor3.Text

                });
            }
            if (Inferior_Campo4.Text != "" && Inferior_Valor4.Text != "")
            {
                FiltrosFacturas.Add(new Filtro
                {
                    Campo = Inferior_Campo4.Text,
                    Valor = Inferior_Valor4.Text

                });
            }

            if (rb_E.IsChecked == true)
            {
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaEgresos);
            }
            else
            {
                dgFacturas.ItemsSource = CrearDataGridFacturas(tablaIngresos);
            }
            timerFacturas.Stop();
        }
        private void busquedaFiltrosBanco(object sender, EventArgs e)
        {
            if (FiltrosBancos == null)
                return;

            FiltrosBancos.Clear();

            if (Superior_Campo1.Text != "" && Superior_Valor1.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = Superior_Campo1.Text,
                    Valor = Superior_Valor1.Text

                });
            }
            if (Superior_Campo2.Text != "" && Superior_Valor2.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = Superior_Campo2.Text,
                    Valor = Superior_Valor2.Text

                });
            }
            if (Superior_Campo3.Text != "" && Superior_Valor3.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = Superior_Campo3.Text,
                    Valor = Superior_Valor3.Text

                });
            }
            if (Superior_Campo4.Text != "" && Superior_Valor4.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = Superior_Campo4.Text,
                    Valor = Superior_Valor4.Text

                });
            }

            if (rb_E.IsChecked == true)
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, false);
            }
            else
            {
                dgBancos.ItemsSource = CrearDataGridBancos(tablaBancos, true);
            }
            timerBancos.Stop();
        }
        private List<ClassFacturas> busquedaComplementos(List<ClassFacturas> facturas, string nombretabla)
        {
            ExcelDatosTabla tbl;
            if(nombretabla == "TblFactEgresos")
            {
                tbl = tablaEgresos;
            }
            else
            {
                tbl = tablaIngresos;
            }
            
            string UUID;
            int n = facturas.Count;
            for(int i = 0; i < n; i++)
            {
                if (!string.IsNullOrEmpty(facturas[i].RelacionREP) &&
                    facturas[i].RelacionREP.StartsWith("CFDI: "))
                {
                    UUID = facturas[i].RelacionREP.Substring(6, 36);
                    int fila = 1;
                    while (fila <= tbl.NumFilas)
                    {
                        if(UUID == tbl.GetValor(fila, "Folio UUID CFDI").ToString())
                        {
                            facturas.Add(new ClassFacturas
                            {
                                filaTabla = fila,
                            });
                        }
                        fila ++;
                    }
                }
            }
            return facturas;
        }
    }
}
