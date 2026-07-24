using Contabilidad.Modelos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace Contabilidad.Formularios.ConfiguracionLecturaFact
{
    /// <summary>
    /// Lógica de interacción para GeneralUserControl1.xaml
    /// </summary>
    public partial class GeneralUserControl1 : UserControl
    {
        private Usuario _usuario;
        private static readonly Regex RegexHex =
            new Regex("^#[0-9A-Fa-f]{6}$");
        public GeneralUserControl1(Usuario usuario)
        {
            InitializeComponent();
            _usuario = usuario;

            InicializarColores();
        }

        private void Color_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox txt = (TextBox)sender;

            if (!RegexHex.IsMatch(txt.Text))
                return;

            Border muestra = null;

            if (txt == txtBoxColorP)
                muestra = muestraColorP;
            else if (txt == txtBoxColorS)
                muestra = muestraColorS;
            else if (txt == txtBoxColorA)
                muestra = muestraColorA;
            else if (txt == txtBoxColorE)
                muestra = muestraColorE;

            if (muestra != null)
            {
                try
                {
                    muestra.Background =
                    (Brush)new BrushConverter().ConvertFromString(txt.Text);
                }
                catch (Exception ex) 
                {
                    MessageBox.Show(ex.ToString());
                }
                
            }
        }
        private void InicializarColores()
        {
            txtBoxColorP.Text = _usuario.ColorPrincipal;
            txtBoxColorS.Text = _usuario.ColorSecundario;
            txtBoxColorA.Text = _usuario.ColorAcento;
            txtBoxColorE.Text = _usuario.ColorElevado;
        }
        private void GuardarColores(object sender, RoutedEventArgs e)
        {
            _usuario.ColorPrincipal = txtBoxColorP.Text;
            _usuario.ColorSecundario = txtBoxColorS.Text;
            _usuario.ColorAcento = txtBoxColorA.Text;
            _usuario.ColorElevado = txtBoxColorE.Text;

            var ventana = Window.GetWindow(this) as ConfiguracionLecturaFacturas;
            ventana.AplicarTema();
        }
    }
}
