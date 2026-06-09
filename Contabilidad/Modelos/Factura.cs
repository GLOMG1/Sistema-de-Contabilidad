using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contabilidad.Modelos
{
    internal class Factura
    {
        public NodoComprobante Comprobante { get; set; }
        public NodoEmisor Emisor { get; set; }
        public NodoReceptor Receptor { get; set; }
        public NodoCfdiRelacionado CfdiRelacionado { get; set; }
        public List<NodoConcepto> Conceptos { get; set; }
        public NodoImpuestosTotales Impuestos { get; set; }
        public NodoComplemento Complemento { get; set; }
        public NodoPago Pago { get; set; }

        public class NodoComprobante
        {
            public string Serie { get; set; }
            public string Folio { get; set; }
            public string Fecha { get; set; }
            public string FormaPago { get; set; }
            public string MetodoPago { get; set; }
            public string TipoComprobante { get; set; }
            public string CondicionesPago { get; set; }
            public decimal? SubTotal { get; set; }
            public decimal? Total { get; set; }
            public string Moneda { get; set; }
            public string Version { get; set; }
            public string Descuento { get; set; }
        }
        public class NodoEmisor
        {
            public string Rfc { get; set; }
            public string Nombre { get; set; }
            public string RegimenFiscal { get; set; }
        }
        public class NodoReceptor
        {
            public string Rfc { get; set; }
            public string Nombre { get; set; }
            public string DomicilioFiscalReceptor { get; set; }
            public string RegimenFiscal { get; set; }
            public string UsoCFDI { get; set; }
        }
        public class NodoCfdiRelacionado
        {
            public string TipoRelacion { get; set; }
            public List<string> UUIDs { get; set; }
        }
        public class NodoComplemento
        {
            public string FechaTimbrado { get; set; }
            public string UUID { get; set; }
        }
        public class NodoConcepto
        {
            public string ClaveProdServ { get; set; }
            public string Cantidad { get; set; }
            public string ClaveUnidad { get; set; }
            public string Unidad { get; set; }
            public string Descripcion { get; set; }
            public string ValorUnitario { get; set; }
            public string Importe { get; set; }
            public string ObjetoImp { get; set; }

        }
        public class NodoImpuestosTotales
        {
            public decimal? ToTotalImpuestosRetenidos { get; set; }
            public decimal? TotalImpuestosTrasladados { get; set; }
            public List<NodoTraslado> Traladados { get; set; }
            public List<NodoRetencion> Retenidos { get; set; }
        }
        public class NodoTraslado
        {
            public decimal? Base { get; set; }
            public string Impuesto { get; set; }   // 002 = IVA, 003 = IEPS
            public string TipoFactor { get; set; } // Tasa, Cuota, Exento
            public string TasaOCuota { get; set; }
            public decimal? Importe { get; set; }
        }
        public class NodoRetencion
        {
            public decimal? Base { get; set; }
            public string Impuesto { get; set; }
            public string TipoFactor { get; set; }
            public string TasaOCuota { get; set; }
            public decimal? Importe { get; set; }
        }
        public class NodoPago
        {
            public string FechaPago { get; set; }
            public string FormaDePagoP { get; set;}
            public decimal? Monto { get; set; }
            public List<NodoDocumentoRelacionado> DocRelacionado { get; set; }

        }
        public class NodoDocumentoRelacionado
        {
            public string UUID { get; set; }
            public decimal? ImpPagado { get; set; }
            public string NumParcialidad { get; set; }
        }
    }
}
