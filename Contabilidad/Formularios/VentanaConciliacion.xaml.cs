using Contabilidad.Modelos;
using Contabilidad.Servicios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace Contabilidad.Formularios
{
    public class Filtro
    {
        public string Campo { get; set;  }
        public string Valor {  get; set; }
    }
    /// <summary>
    /// Lógica de interacción para VentanaConciliacion.xaml
    /// </summary>
    public partial class VentanaConciliacion : Window
    {
        /// Elementos de uso dentro de la venta
        Excel_Function excel = new Excel_Function();
        public List<ClassBancos> tblBancos { get; set; }
        public List<String> EncabezadosBancos { get; set;  }
        public List<ClassFacturas> tblFacturasEgreso {  get; set; }
        public List<ClassFacturas> tblFacturasIngreso { get; set; }
        public List<String> EncabezadosFacturas { get; set; }
        public List<Filtro> FiltrosBancos { get; set; }
        public List<Filtro> FiltrosFacturas {  get; set; }

        /// Metodos esenciales dentro de la ventana
        public VentanaConciliacion()
        {
            InitializeComponent();

            EncabezadosBancos = excel.Encabezados("EstadoCuenta");
            EncabezadosFacturas = excel.Encabezados("TblFactEgresos");
            tblFacturasEgreso = excel.ObtenerFacturas("TblFactEgresos");
            FiltrosBancos = new List<Filtro>();
            FiltrosFacturas = new List<Filtro>();
            dgFacturas.ItemsSource = tblFacturasEgreso;
        }
        public void MostarDataGrid_Cheked(object sender, RoutedEventArgs e)
        {
            if (rb_E.IsChecked == true)
            {
                List<ClassBancos> bancos = excel.ObtenerBancos(true, FiltrosBancos);
                dgBancos.ItemsSource = bancos;
            }
            if (rb_I.IsChecked == true)
            {
                List<ClassBancos> bancos = excel.ObtenerBancos(false, FiltrosBancos);
                dgBancos.ItemsSource = bancos;
            }
        }
        public void CambiosFiltros(object sender, KeyEventArgs e)
        {
            List<string> ListaFiltrada;
            ComboBox combo = (ComboBox)sender;
            string texto = combo.Text.ToLower();

            ListaFiltrada = EncabezadosBancos
                .Where(x => x.ToLower().Contains(texto))
                .ToList();

            combo.ItemsSource = ListaFiltrada;
            combo.IsDropDownOpen = ListaFiltrada.Any();
        }
        public void CambiosValores(object sender, TextChangedEventArgs e)
        {
            FiltrosBancos.Clear();
            
            if(tb_Campo1.Text != "" && tb_Valor1.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = tb_Campo1.Text,
                    Valor = tb_Valor1.Text,
                });
            }
            if (tb_Campo2.Text != "" && tb_Valor2.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = tb_Campo2.Text,
                    Valor = tb_Valor2.Text,
                });
            }
            if (tb_Campo3.Text != "" && tb_Valor3.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = tb_Campo3.Text,
                    Valor = tb_Valor3.Text,
                });
            }
            if (tb_Campo4.Text != "" && tb_Valor4.Text != "")
            {
                FiltrosBancos.Add(new Filtro
                {
                    Campo = tb_Campo4.Text,
                    Valor = tb_Valor4.Text,
                });
            }

            if (rb_E.IsChecked == true)
            {
                List<ClassBancos> bancos = excel.ObtenerBancos(true, FiltrosBancos);
                dgBancos.ItemsSource = bancos;
            }
            if (rb_I.IsChecked == true)
            {
                List<ClassBancos> bancos = excel.ObtenerBancos(false, FiltrosBancos);
                dgBancos.ItemsSource = bancos;
            }
        }

        /// Funciones para Metodos

    }
}
