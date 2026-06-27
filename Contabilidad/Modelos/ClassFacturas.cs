using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contabilidad.Modelos
{
    public class ClassFacturas
    {
        // Identificadores y control
        public int filaTabla {  get; set; }
        public int IdRegistroBD { get; set; }
        public string Fiscal { get; set; }
        public string Sublibro { get; set; }

        // Datos del comprobante
        public string SerieFolioInterno { get; set; }
        public string RfcEmisor { get; set; }
        public string NombreEmisor { get; set; }
        public DateTime FechaCFDI { get; set; }
        public string TipoCFDI { get; set; }
        public string UsoCFDI { get; set; }
        public string FolioUUID { get; set; }
        public string FormaPago { get; set; }
        public string MetodoPago { get; set; }
        public string Concepto { get; set; }
        public string IcaRegistro { get; set; }

        // Importes fiscales
        public decimal SubtotalCFDI { get; set; }
        public decimal IvaCFDI { get; set; }
        public decimal IepsCFDI { get; set; }
        public decimal IsrRetenido { get; set; }
        public decimal IvaRetenido { get; set; }
        public decimal TotalCFDI { get; set; }

        // Pago
        public decimal ProporcionPago { get; set; }
        public DateTime? FechaPago { get; set; }

        // Contabilidad
        public string CuentaContProveedor { get; set; }
        public string RutaArchivoPC { get; set; }

        // Diferencias
        public decimal DiferenciaTotal { get; set; }
        public decimal DiferenciaIVA { get; set; }

        // Bancos
        public int? IdRegistroBancos { get; set; }
        public DateTime? FechaBancos { get; set; }
        public string FolioBancos { get; set; }

        // Estatus y cuenta bancaria
        public string Estatus { get; set; }
        public string Banco { get; set; }
        public string NumeroCuenta { get; set; }
    }
}
