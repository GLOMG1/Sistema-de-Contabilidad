using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Contabilidad.Modelos;
using static Contabilidad.Modelos.Factura;

namespace Contabilidad.Servicios
{
    internal class Herramientas
    {
        public object[] Añadir;

        public int IdFolioUUID;
        public int TipoFact;
        private Dictionary<string, int> columna;

        private void crearDictionary(Excel.ListObject tbl)
        {
            columna = new Dictionary<string, int>();

            foreach (Excel.ListColumn col in tbl.ListColumns)
            {
                columna.Add(col.Name, col.Index);
            }
        }

        public object[] ContructorFila(Factura factura, Excel.ListObject tbl)
        {
            crearDictionary(tbl);
            IdFolioUUID = IdColumna("Folio UUID CFDI");
            TipoFact = IdColumna("Tipo CFDI");
            object[] fila = new object[columna.Count];

            fila[columna["Serie Folio Interno CFDI"]-1]     = $"F: {factura.Comprobante.Serie}{factura.Comprobante.Folio}";
            fila[columna["RFC Emisor CFDI"] - 1]            = factura.Emisor.Rfc;
            fila[columna["Nombre Emisor CFDI"] - 1]         = factura.Emisor.Nombre;
            fila[columna["RFC Receptor CFDI"] - 1]          = factura.Receptor.Rfc;
            fila[columna["Nombre Receptor CFDI"] - 1]       = factura.Receptor.Nombre;
            fila[columna["Fecha CFDI"] - 1]                 = factura.Comprobante.Fecha;
            fila[columna["Tipo CFDI"] - 1]                  = factura.Comprobante.TipoComprobante;
            fila[columna["Uso CFDI"] - 1]                   = factura.Receptor.UsoCFDI;
            fila[columna["Version CFDI"] - 1]               = factura.Comprobante.Version;
            fila[columna["Folio UUID CFDI"] - 1]            = factura.Complemento.UUID;
            fila[columna["Moneda CFDI"] - 1]                = factura.Comprobante.Moneda;
            fila[columna["Tipo Cambio CFDI"] - 1]           = "01";
            fila[columna["Forma Pago CFDI"] - 1]            = factura.Comprobante.FormaPago;
            fila[columna["Metodo Pago CFDI"] - 1]           = factura.Comprobante.MetodoPago;
            fila[columna["Condicion Pago CFDI"] - 1]        = factura.Comprobante.CondicionesPago;
            fila[columna["Concepto CFDI"] - 1]              = string.Join(" - ", factura.Conceptos.Select(c => c.Descripcion));

            switch (factura.Comprobante.TipoComprobante)
            {
                case "I":
                    if (factura.Impuestos == null)
                    {
                        fila[columna["Gravado 16"] - 1] = 0;
                        fila[columna["Gravado 0"] - 1] = 0;
                        fila[columna["Exento"] - 1] = 0;
                        fila[columna["Sin IVA"] - 1] = factura.Comprobante.SubTotal;
                        fila[columna["Descuento 16"] - 1] = 0;
                        fila[columna["Descuento 0"] - 1] = 0;
                        fila[columna["Descuento Exento"] - 1] = 0;
                        fila[columna["Descuento Sin IVA"] - 1] = 0;
                        fila[columna["SUBTOTAL CFDI"] - 1] = "";
                        fila[columna["IVA CFDI"] - 1] = 0;
                        fila[columna["ISH CFDI"] - 1] = 0;
                        fila[columna["ISR Retenido CFDI"] - 1] = 0;
                        fila[columna["IVA Retenido CFDI"] - 1] = 0;
                        fila[columna["TOTAL CFDI"] - 1] = factura.Comprobante.Total;
                    }
                    else
                    {
                        decimal[] imp = DeterminadorImpuestos(factura);

                        fila[columna["Gravado 16"] - 1] = imp[0];
                        fila[columna["Gravado 0"] - 1] = imp[1];
                        fila[columna["Exento"] - 1] = imp[2];
                        fila[columna["Sin IVA"] - 1] = imp[3];
                        fila[columna["Descuento 16"] - 1] = imp[4];
                        fila[columna["Descuento 0"] - 1] = imp[5];
                        fila[columna["Descuento Exento"] - 1] = imp[6];
                        fila[columna["Descuento Sin IVA"] - 1] = imp[7];
                        fila[columna["SUBTOTAL CFDI"] - 1] = "";
                        fila[columna["IVA CFDI"] - 1] = imp[9];
                        //fila[columna["ISH CFDI"] - 1] = imp[10];
                        fila[columna["ISR Retenido CFDI"] - 1] = imp[11];
                        fila[columna["IVA Retenido CFDI"] - 1] = imp[12];
                        fila[columna["TOTAL CFDI"] - 1] = factura.Comprobante.Total;
                    }
                    break;
                case "P":
                    if (factura.Pago == null)
                    {

                    }
                    else
                    {
                        fila[columna["Fecha Pago"] - 1] = factura.Pago.FechaPago;
                        fila[columna["TOTAL CFDI"] - 1] = factura.Pago.Monto;

                        int i = factura.Pago.DocRelacionado.Count;
                        Añadir = new object[i];
                        i = 0;
                        foreach (NodoDocumentoRelacionado doc in factura.Pago.DocRelacionado)
                        {
                            fila[columna["Concepto CFDI"] - 1] = $"{factura.Pago.FechaPago}; MONTO TOTAL ${factura.Pago.Monto?.ToString("N2")}; NUMERO DE PARCIALIDAD: {doc.NumParcialidad}";
                            fila[columna["Folio UUID CFDI Relacionado REP"] - 1] = doc.UUID;
                            fila[columna["TOTAL CFDI"] - 1] = doc.ImpPagado;

                            Añadir[i] = (object)fila.Clone();
                            i++;
                        }
                    }
                    break;

                default:

                    break;
            }
            fila[columna["SUBTOTAL CFDI"] - 1] = "=SUM(tblEgresos[@[Gravado 16]:[Descuento Sin IVA]])";
            return fila;
        }

        private decimal[] DeterminadorImpuestos(Factura factura)
        {
            decimal[] tblImpuestos = new decimal[13];
            
            decimal importeTrasladado = factura.Impuestos.TotalImpuestosTrasladados ?? 0;
            decimal importeRetenido = factura.Impuestos.ToTotalImpuestosRetenidos ?? 0;
            
            if (factura.Impuestos.Traladados != null)
            {
                foreach (Factura.NodoTraslado impuestoT in factura.Impuestos.Traladados)
                {
                    switch (impuestoT.Impuesto)
                    {
                        case "001":

                            break;
                        case "002":
                            if (impuestoT.TipoFactor == "Exento")
                            {
                                tblImpuestos[2] = impuestoT.Base ?? 0;
                            }
                            else
                            {
                                switch (impuestoT.TasaOCuota)
                                {
                                    case "0.160000":
                                        tblImpuestos[0] = impuestoT.Base ?? 0;
                                        break;
                                    case "0.000000":
                                        tblImpuestos[1] = impuestoT.Base ?? 0;
                                        break;
                                    default:
                                        System.Windows.MessageBox.Show(factura.Complemento.UUID);
                                        break;
                                }
                                tblImpuestos[9] = impuestoT.Importe ?? 0;
                            }
                            break;
                        case "003":

                            break;
                        default:

                            break;
                    }
                }
            }

            if(factura.Impuestos.Retenidos != null)
            {
                foreach (Factura.NodoRetencion impuestoR in factura.Impuestos.Retenidos)
                {
                    switch (impuestoR.Impuesto)
                    {
                        case "001":
                            tblImpuestos[11] = (impuestoR.Importe * -1 ) ?? 0;
                            break;
                        case "002":
                            tblImpuestos[12] = (impuestoR.Importe * -1) ?? 0;
                            break;
                        default:
                            tblImpuestos[7] = (impuestoR.Importe * -1) ?? 0;
                            break;
                    }
                }
            }

            for (int i = 0; i <= 12; i++)
            {
                if (tblImpuestos[i] == 0)
                {
                    tblImpuestos[i] = 0;
                }
            }

            tblImpuestos[3] = (factura.Comprobante.Total ?? 0) - importeTrasladado + importeRetenido - tblImpuestos[0] - tblImpuestos[1] - tblImpuestos[2];

            return tblImpuestos;
        }

        private int IdColumna(string colum)
        {
            return (columna[colum] - 1);
        }
    }
}
