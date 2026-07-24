using Contabilidad.Formularios;
using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Contabilidad.Servicios;
using Excel= Microsoft.Office.Interop.Excel;
using System.Drawing;
using Contabilidad.Modelos;
using Contabilidad.Servicios.SQL;

namespace Contabilidad
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public Ribbon1()
        {
        }

        #region Miembros de IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Contabilidad.Ribbon1.xml");
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opcionesf
        //Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void Refrescar()
        {
            ribbon?.Invalidate();
        }
        public Bitmap GetImage(Office.IRibbonControl control)
        {
            switch(control.Id)
            {
                case "btnIniciar":
                    return Properties.Resources.Usuarios;
                case "btnImportar":
                    return Properties.Resources.Importar;
                case "btnGuardarBD":
                    return Properties.Resources.Guardar;
                case "btnImportarEC":
                    return Properties.Resources.Bancos;
                case "btnConciliar":
                    return Properties.Resources.Conciliar;
                case "btnCatalogoCuentas":
                    return Properties.Resources.Balance;
                case "btnEventoEco":
                    return Properties.Resources.Evento_Economico;
                case "btnDeterminarImp":
                    return Properties.Resources.SAT;
            }
            return null;
        }
        public void btnIniciar(Office.IRibbonControl control)
        {
            var ventana = new Formularios.Principal();
            ventana.ShowDialog();
            
        }
        public void btonAbrirCarpeta(Office.IRibbonControl control)
        {
            Process.Start("explorer.exe", $"\\\\CONTA\\Users\\ERIKA\\Documents\\CLIENTES DESPACHO\\CATAR\\2026");
        }
        public void btnConciliar(Office.IRibbonControl control)
        {
            var ventana = new Formularios.VentanaConciliacion();
            ventana.ShowDialog();
        }
        public void btnConfiguraciones(Office.IRibbonControl control)
        {
            var ventana = new Formularios.ConfiguracionLecturaFacturas();
            ventana.ShowDialog();
        }
        public void btnImportar(Office.IRibbonControl control)
        {
            Excel_Function x = new Excel_Function();
            x.CargarDatos();
        }
        public void btnCrearEjercicio(Office.IRibbonControl control)
        {
            gestor_Archivos.CrearCarpetas();
        }
        public void btonPruebas(Office.IRibbonControl control)
        {
            Excel_Function x = new Excel_Function();
            x.FuncionDePrueba();
        }
        public bool ValidarLibro_GetEnabled(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.VerificacionLibro();
        }
        #endregion

        #region Asistentes

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
