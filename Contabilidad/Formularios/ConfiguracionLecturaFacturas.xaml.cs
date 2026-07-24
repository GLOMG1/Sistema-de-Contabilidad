using Contabilidad.Formularios.ConfiguracionLecturaFact;
using Contabilidad.Modelos;
using Contabilidad.Servicios.SQL;
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
    /// <summary>
    /// Lógica de interacción para ConfiguracionLecturaFacturas.xaml
    /// </summary>
    public partial class ConfiguracionLecturaFacturas : Window
    {
        public Usuario UsuarioActivo { get; set; }
        private Dictionary<string, UserControl> _vistas;
        public ConfiguracionLecturaFacturas()
        {
            InitializeComponent();

            UsuarioActivo = InicioRepositorio.UsuarioActivo();
            AplicarTema();

            _vistas = new Dictionary<string, UserControl>
            {
                ["Vistas"] = new VistasUserControl(),
                ["Reglas"] = new ReglasUserControl(),
                ["Catalogos"] = new CatalogosUserControl(),
                ["General"] = new GeneralUserControl1(UsuarioActivo),
            };
        }
        private void MenuOpcion_Checked(object sender, RoutedEventArgs e)
        {
            var tag = (sender as RadioButton)?.Tag as string;
            if (tag != null && _vistas.TryGetValue(tag, out var vista))
                PanelContenido.Content = vista;
        }
        public void AplicarTema()
        {
            if (UsuarioActivo == null)
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
