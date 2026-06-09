using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Xml.Linq;
using Contabilidad.Modelos;


namespace Contabilidad.Servicios
{
    public class RutaDeTrabajo
    {
        public string rutaPrincipal { get; set;}
        public string rutaCliente { get; set;}
    }

    public class gestor_Archivos
    {
        static public string AbrirInfoClientes()
        {
            string ruta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "Configuracion.json");

            string json = File.ReadAllText(ruta);

            return json;
        }
        private string RutaArchivos()
        {
            using (var ventana = new CommonOpenFileDialog())
            {
                ventana.IsFolderPicker = true;
                ventana.Title = "Seleccione la ruta de los CFDIs";

                if (ventana.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    return ventana.FileName;
                }
            }
            return null;
        }
        public IEnumerable<string> ArchivosXML()
        {
            string direccion = RutaArchivos();
            if (direccion == null)
                return null;

            return Directory.EnumerateFiles(
                direccion,
                "*.xml",
                SearchOption.AllDirectories
                );
        }
        static public void CrearCarpetas()
        {
            string direccion;
            using (var ventana = new CommonOpenFileDialog())
            {
                ventana.IsFolderPicker = true;
                ventana.Title = "Seleccione la ruta de los CFDIs";

                if (ventana.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    direccion = ventana.FileName;
                }
                else
                {
                    direccion = null;
                }
            }
            if (direccion != null)
            {
                string[] arbolCarpetas = { 
                    "\\Ejercicio 2026" , 
                    "\\Ejercicio 2026\\Expediente Fiscal", 
                    "\\Ejercicio 2026\\Facturas Recibidas", 
                    "\\Ejercicio 2026\\Facturas Emitidas", 
                    "\\Ejercicio 2026\\Nomina", 
                    "\\Ejercicio 2026\\Bancos", 
                    "\\Ejercicio 2026\\Expediente Contable"
                };
                string[] meses = {
                    "\\01 - Enero",
                    "\\02 - Febrero",
                    "\\03 - Marzo",
                    "\\04 - Abril",
                    "\\05 - Mayo",
                    "\\06 - Junio",
                    "\\07 - Julio",
                    "\\08 - Agosto",
                    "\\09 - Septiembre",
                    "\\10 - Octubre",
                    "\\11 - Noviembre",
                    "\\12 - Diciembre"
                };
                
                for(int i = 0; i <= 6; i++)
                {
                    if(i != 0)
                    {
                        foreach(string mes in meses)
                        {
                            Directory.CreateDirectory(direccion + arbolCarpetas[i]+mes);
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(direccion + arbolCarpetas[i]);
                    }
                }

            }
        }

    }
    internal class xml_Function
    {
        private bool validador(string xml)
        {
            XDocument DocXML = XDocument.Load(xml);
            XNamespace cfdi = "http://www.sat.gob.mx/cfd/4";

            XElement comprobante = DocXML.Element(cfdi + "Comprobante");

            if (comprobante != null)
                return true;
            else
                return false;
        }
        public Factura leerArchivos(string xml)
        {
            if (!validador(xml))
                return null;

            XDocument CFDI = XDocument.Load(xml);
            XNamespace Ncfdi = "http://www.sat.gob.mx/cfd/4";
            XNamespace Npago = "http://www.sat.gob.mx/Pagos20";
            XNamespace nTdf = "http://www.sat.gob.mx/TimbreFiscalDigital";

            XElement comprobante = CFDI.Element(Ncfdi + "Comprobante");
            XElement emisor = comprobante.Element(Ncfdi + "Emisor");
            XElement receptor = comprobante.Element(Ncfdi + "Receptor");
            XElement conceptos = comprobante?.Element(Ncfdi + "Conceptos");

            XElement complemento = comprobante.Element(Ncfdi + "Complemento");
            XElement Pagos = complemento.Element(Npago + "Pagos");

            XElement timbreFiscal = complemento.Element(nTdf + "TimbreFiscalDigital");

            XElement impuestos = comprobante.Element(Ncfdi + "Impuestos");

            Factura modelo = new Factura
            {
                Comprobante = new Factura.NodoComprobante
                {
                    TipoComprobante = (string)comprobante.Attribute("TipoDeComprobante").Value,
                    Version = (string)comprobante.Attribute("Version")?.Value,
                    Serie = (string)comprobante.Attribute("Serie")?.Value,
                    Folio = (string)comprobante.Attribute("Folio")?.Value,
                    Fecha = (string)comprobante.Attribute("Fecha").Value,
                    FormaPago = (string)comprobante.Attribute("FormaPago")?.Value,
                    MetodoPago = (string)comprobante.Attribute("MetodoPago")?.Value,
                    CondicionesPago = (string)comprobante.Attribute("CondicionesDePago")?.Value,
                    SubTotal = (decimal?)comprobante.Attribute("SubTotal"),
                    Descuento = (string)comprobante.Attribute("Descuento")?.Value,
                    Total = (decimal?)comprobante.Attribute("Total"),
                    Moneda = (string)comprobante.Attribute("Moneda")?.Value
                },
                Emisor = new Factura.NodoEmisor
                {
                    Rfc = (string)emisor.Attribute("Rfc"),
                    Nombre = (string)emisor.Attribute("Nombre"),
                    RegimenFiscal = (string)emisor.Attribute("RegimenFiscal")
                },
                Receptor = new Factura.NodoReceptor
                {
                    Rfc = (string)receptor.Attribute("Rfc"),
                    Nombre = (string)receptor.Attribute("Nombre"),
                    RegimenFiscal = (string)receptor.Attribute("RegimenFiscal"),
                    DomicilioFiscalReceptor = (string)receptor.Attribute("DomicilioFiscalReceptor"),
                    UsoCFDI = (string)receptor.Attribute("UsoCFDI")
                },
                Complemento = new Factura.NodoComplemento
                {
                    FechaTimbrado = (string)timbreFiscal.Attribute("FechaTimbrado"),
                    UUID = (string)timbreFiscal.Attribute("UUID")
                },

                Conceptos = conceptos?
                    .Elements(Ncfdi + "Concepto")
                    .Select(c => new Factura.NodoConcepto
                    {
                        ClaveProdServ = (string)c.Attribute("ClaveProdServ"),
                        Descripcion = (string)c.Attribute("Descripcion")
                    }).ToList(),

                Impuestos = impuestos == null ? null : new Factura.NodoImpuestosTotales
                {
                    TotalImpuestosTrasladados = (decimal?)impuestos.Attribute("TotalImpuestosTrasladados"),
                    ToTotalImpuestosRetenidos = (decimal?)impuestos.Attribute("TotalImpuestosRetenidos"),
                    Traladados = impuestos.Element(Ncfdi + "Traslados")?
                        .Elements(Ncfdi + "Traslado")
                        .Select(t => new Factura.NodoTraslado
                        {
                            Base = (decimal?)t.Attribute("Base"),
                            Impuesto = (string)t.Attribute("Impuesto"),
                            TipoFactor = (string)t.Attribute("TipoFactor"),
                            TasaOCuota = (string)t.Attribute("TasaOCuota"),
                            Importe = (decimal?)t.Attribute("Importe")
                        }).ToList(),
                    Retenidos = impuestos.Element(Ncfdi + "Retenciones")?
                        .Elements(Ncfdi + "Retencion")
                        .Select(r => new Factura.NodoRetencion
                        {
                            Base = (decimal?)r.Attribute("Base"),
                            Impuesto = (string)r.Attribute("Impuesto"),
                            TipoFactor = (string)r.Attribute("TipoFactor"),
                            TasaOCuota = (string)r.Attribute("TasaOCuota"),
                            Importe = (decimal?)r.Attribute("Importe")
                        }).ToList()
                }
            };

            if (modelo.Comprobante.TipoComprobante == "P")
                modelo.Pago = Reps(Pagos);

            return modelo;
        }
        private Factura.NodoPago Reps(XElement TipoP)
        {
            XNamespace Npago = "http://www.sat.gob.mx/Pagos20";
            XElement NodoPago = TipoP.Element(Npago + "Pago");

            Factura.NodoPago ObjetoFactura = new Factura.NodoPago();

            ObjetoFactura.FechaPago = (string)NodoPago.Attribute("FechaPago")?.Value;
            ObjetoFactura.Monto = (decimal?)NodoPago.Attribute("Monto");
            ObjetoFactura.FormaDePagoP = (string)NodoPago.Attribute("FormaDePagoP");
            ObjetoFactura.DocRelacionado = NodoPago
                .Elements(Npago + "DoctoRelacionado")
                .Select(p => new Factura.NodoDocumentoRelacionado 
                {
                    UUID = (string)p.Attribute("IdDocumento"),
                    ImpPagado = (decimal?)p.Attribute("ImpPagado"),
                    NumParcialidad = (string)p.Attribute("NumParcialidad")
                })
                .ToList();


            return ObjetoFactura;
        }
    }
}
