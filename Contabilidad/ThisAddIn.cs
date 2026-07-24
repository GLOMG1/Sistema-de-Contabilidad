using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Contabilidad.Servicios;
using Contabilidad.Servicios.SQL;
using System.Windows;
using Contabilidad.Modelos;

namespace Contabilidad
{
    public partial class ThisAddIn
    {
        private Ribbon1 RibbonConta;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonConta = new Ribbon1();
            return RibbonConta;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += Application_WorkbookActivate;
            
            ConexionDb.Iniciar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WorkbookActivate(Excel.Workbook wb)
        {
            RibbonConta?.Refrescar();
        }
        public bool VerificacionLibro()
        {
            try
            {
                Excel.Workbook wb = this.Application.ActiveWorkbook;
                if (wb == null)
                    return false;

                bool Bancos = false;
                bool BaseEgresos = false;
                bool BaseIngresos = false;

                foreach (Excel.Worksheet hoja in wb.Sheets)
                {
                    if (hoja.Name.Equals("Bancos", StringComparison.OrdinalIgnoreCase))
                        Bancos = true;

                    if (hoja.Name.Equals("BaseEgresos", StringComparison.OrdinalIgnoreCase))
                        BaseEgresos = true;

                    if (hoja.Name.Equals("BaseIngresos", StringComparison.OrdinalIgnoreCase))
                        BaseIngresos = true;
                }

                return Bancos && BaseEgresos && BaseIngresos;

            }
            catch (System.Exception)
            {
                return false;
            }
        }
        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
