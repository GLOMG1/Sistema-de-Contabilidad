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
using Contabilidad.Servicios;
using System.Collections.ObjectModel;

namespace Contabilidad.Formularios
{
    public partial class Principal : Window
    {
        public ObservableCollection<Contribuyentes> listClientes { get; set; }
        public Contribuyentes clienteSeleccionado { get; set; }
        public Principal()
        {
            InitializeComponent();
            DbContabilidad.Iniciar();

            var lista = DbContabilidad.Consultar();
            listClientes = new ObservableCollection<Contribuyentes>(lista);

            this.DataContext = this;
        }
        private void Nuevo_Click(object sender, RoutedEventArgs e)
        {

            if(panelClientes.Visibility == Visibility.Visible)
            {
                panelClientes.Visibility = Visibility.Collapsed;
                panelAltaClientes.Visibility = Visibility.Visible;
            }
            else{
                panelClientes.Visibility = Visibility.Visible;
                panelAltaClientes.Visibility = Visibility.Collapsed;
            }

        }

        private void AgregarCliente_Click(object sender, RoutedEventArgs e)
        {
            string nombre = txtNombre.Text;
            string rfc = txtRfc.Text;
            int regimen = txtRegimen.LineCount;

            if (string.IsNullOrEmpty(nombre) || string.IsNullOrEmpty(rfc))
            {
                MessageBox.Show("Ponte trucha bb");
            }
            else
            {
                DbContabilidad.Insertar(nombre, rfc, regimen);
            }
        }

        private void Leer_Click(object sender, RoutedEventArgs e)
        {
            var lista = DbContabilidad.Consultar();

            foreach (var c in lista)
            {
                MessageBox.Show($"{c.Id} - {c.Nombre} - {c.Rfc}");
            }
        }

        private void Iniciar_DobleClick(object sender, MouseButtonEventArgs e)
        {
            Excel_Function x = new Excel_Function();
            x.CrearTabla(clienteSeleccionado);
            this.Close();
        }
    }
}
