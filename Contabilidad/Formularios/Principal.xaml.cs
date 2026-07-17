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
using Contabilidad.Servicios.SQL;
using Contabilidad.Modelos;

namespace Contabilidad.Formularios
{
    public partial class Principal : Window
    {
        public ObservableCollection<Contribuyentes> listClientes { get; set; }
        public Contribuyentes clienteSeleccionado { get; set; }
        private Usuario UsuarioActivo { get; set; }
        private InicioRepositorio _db;
        public Principal()
        {
            InitializeComponent();

            _db = new InicioRepositorio();

            UsuarioActivo = _db.UsuarioActivo();
            AplicarTema(UsuarioActivo);

            var lista = _db.Consultar();
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
                _db.Insertar(nombre, rfc, regimen);
            }
        }
        private void Leer_Click(object sender, RoutedEventArgs e)
        {
            var lista = _db.Consultar();

            foreach (var c in lista)
            {
                MessageBox.Show($"{c.Id} - {c.Nombre} - {c.Rfc}");
            }
        }
        private void EliminarRegistro_Click(object sender, RoutedEventArgs e)
        {
            if(clienteSeleccionado == null)
            {
                MessageBox.Show("Se necesita tener un contribuente seleccionado");
                return;
            }

            MessageBoxResult resultado = MessageBox.Show(
                string.Format("¿Estas seguro de querer eliminar a {0}?", clienteSeleccionado.Nombre),
                "Confirmar eliminacion",
                MessageBoxButton.YesNo,
                MessageBoxImage.Exclamation);

            if(resultado == MessageBoxResult.Yes)
                _db.Eliminar(clienteSeleccionado.Id);
        }
        private void Iniciar_DobleClick(object sender, MouseButtonEventArgs e)
        {
            Excel_Function x = new Excel_Function();
            x.CrearTabla(clienteSeleccionado);
            this.Close();
        }
        private void AplicarTema(Usuario us)
        {
            if (us == null)
                return;

            this.Resources["ColorPrincipal"] = 
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(UsuarioActivo.ColorPrincipal));
            this.Resources["ColorSecundario"] =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(UsuarioActivo.ColorSecundario));
            this.Resources["ColorAcento"] =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(UsuarioActivo.ColorAcento));
            this.Resources["ColorElevado"] =
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(UsuarioActivo.ColorElevado));

        }
    }
}
