using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Contabilidad.Formularios;
using Contabilidad.Modelos;

namespace Contabilidad.Servicios.SQL
{
    public  class FacturasRepositorio
    {
        private List<ClassFacturas> FacturasEnExcel;
        private List<ClassFacturas> FacturasEnBaseDatos;
        public FacturasRepositorio(ExcelDatosTabla DatosExcel)
        {
            FacturasEnExcel = CrearListaFacturasExcel(DatosExcel);
        }
        private void ConsultarBD()
        {

        }
        public void InsertarFacturas()
        {
            /* Recordatorio de la Base de datos
                CREATE TABLE IF NOT EXISTS Facturas(
                    Id_Facturas INTEGER PRIMARY KEY AUTOINCREMENT
                                              NOT NULL,
                    Clasificacion TEXT    NOT NULL,
                    Fiscal TEXT,
                    Estado TEXT    NOT NULL,
                    AplicaRegistro TEXT,
                    SerieFolioInterno TEXT,
                    RfcEmisor TEXT    NOT NULL,
                    NombreEmisor TEXT    NOT NULL,
                    RfcReceptor TEXT    NOT NULL,
                    NombreReceptor TEXT    NOT NULL,
                    Fecha TEXT    NOT NULL,
                    Tipo TEXT    NOT NULL,
                    Uso TEXT    NOT NULL,
                    Version TEXT    NOT NULL,
                    FolioUuid TEXT    NOT NULL
                                              UNIQUE,
                    Moneda TEXT    NOT NULL,
                    TipoCambio NUMERIC NOT NULL,
                    FormaPago TEXT    NOT NULL,
                    MetodoPago TEXT    NOT NULL,
                    CondicionPago TEXT    NOT NULL,
                    SubTotal NUMERIC,
                    Descuento NUMERIC,
                    Total NUMERIC,
                    LugarExpedicion TEXT,
                    Exportacion INTEGER,
                    FechaTimbrado TEXT    NOT NULL,
                    PorcionPago NUMERIC
                );*/
            Excel_Function x = new Excel_Function();
            using (var connection = EmpresaDb.ObtenerConexion())
            {
                using(var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var command = connection.CreateCommand())
                        {
                            command.Transaction = transaction;
                            command.CommandText =
                                @"
                                INSERT INTO Facturas(Clasificacion, Fiscal, Estado,
                                    SerieFolioInterno, RfcEmisor, NombreEmisor, RfcReceptor, NombreReceptor,
                                    Fecha, Tipo, Uso, Version, FolioUuid, Moneda, TipoCambio, FormaPago,
                                    MetodoPago, CondicionPago, Total)
                                VALUES  (
                                    @Clasificacion,
                                    @Fiscal,
                                    'ACUMULABLE',
                                    @SerieFolioInterno,
                                    @RfcEmisor,
                                    @NombreEmisor,
                                    @RfcReceptor,
                                    @NombreReceptor,
                                    @Fecha,
                                    @Tipo,
                                    @Uso,
                                    @Version,
                                    @FolioUuid,
                                    @Moneda,
                                    @TipoCambio,
                                    @FormaPago,
                                    @MetodoPago,
                                    @CondicionPago,
                                    @Total
                                )
                                ";
                            command.Parameters.AddWithValue("@Clasificacion", DBNull.Value);
                            command.Parameters.AddWithValue("@Fiscal", DBNull.Value);
                            command.Parameters.AddWithValue("@SerieFolioInterno", DBNull.Value);
                            command.Parameters.AddWithValue("@RfcEmisor", DBNull.Value);
                            command.Parameters.AddWithValue("@NombreEmisor", DBNull.Value);
                            command.Parameters.AddWithValue("@RfcReceptor", DBNull.Value);
                            command.Parameters.AddWithValue("@NombreReceptor", DBNull.Value);
                            command.Parameters.AddWithValue("@Fecha", DBNull.Value);
                            command.Parameters.AddWithValue("@Tipo", DBNull.Value);
                            command.Parameters.AddWithValue("@Uso", DBNull.Value);
                            command.Parameters.AddWithValue("@Version", DBNull.Value);
                            command.Parameters.AddWithValue("@FolioUuid", DBNull.Value);
                            command.Parameters.AddWithValue("@Moneda", DBNull.Value);
                            command.Parameters.AddWithValue("@TipoCambio", DBNull.Value);
                            command.Parameters.AddWithValue("@FormaPago", DBNull.Value);
                            command.Parameters.AddWithValue("@MetodoPago", DBNull.Value);
                            command.Parameters.AddWithValue("@CondicionPago", DBNull.Value);
                            command.Parameters.AddWithValue("@Total", DBNull.Value);

                            foreach(ClassFacturas factura in FacturasEnExcel)
                            {
                                x.EscribirDatos("TblFactEgresos", factura.filaTabla, "Sublibro", "ok");
                                command.Parameters["@Clasificacion"].Value = factura.Clasificacion;
                                command.Parameters["@Fiscal"].Value = factura.Fiscal;
                                command.Parameters["@SerieFolioInterno"].Value = factura.SerieFolioInterno;
                                command.Parameters["@RfcEmisor"].Value = factura.RfcEmisor;
                                command.Parameters["@NombreEmisor"].Value = factura.NombreEmisor;
                                command.Parameters["@RfcReceptor"].Value = factura.RfcReceptor;
                                command.Parameters["@NombreReceptor"].Value = factura.NombreReceptor;
                                command.Parameters["@Fecha"].Value = factura.FechaCFDI.ToString("yyyy-MM-dd HH:mm:ss");
                                command.Parameters["@Tipo"].Value = factura.TipoCFDI;
                                command.Parameters["@Uso"].Value = factura.UsoCFDI;
                                command.Parameters["@Version"].Value = factura.Version;
                                command.Parameters["@FolioUuid"].Value = factura.FolioUUID;
                                command.Parameters["@Moneda"].Value = factura.Moneda;
                                command.Parameters["@TipoCambio"].Value = factura.TipoCambio;
                                command.Parameters["@FormaPago"].Value = factura.FormaPago;
                                command.Parameters["@MetodoPago"].Value = factura.MetodoPago;
                                command.Parameters["@CondicionPago"].Value = factura.CondicionPago;
                                command.Parameters["@Total"].Value = factura.TotalCFDI;

                                command.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }
        private List<ClassFacturas> CrearListaFacturasExcel(ExcelDatosTabla tabla)
        {
            List<ClassFacturas> tblFiltrada = new List<ClassFacturas>();
            for (int fila = 1; fila <= tabla.NumFilas; fila++)
            {
                tblFiltrada.Add(new ClassFacturas
                {
                    filaTabla = fila,
                    IdRegistroBD = Convert.ToInt32(tabla.GetValor(fila, "ID Registro BD") ?? 0),
                    Clasificacion = tabla.GetValor(fila, "Clasificacion")?.ToString(),
                    FechaCFDI = DateTime.FromOADate(Convert.ToDouble(tabla.GetValor(fila, "Fecha CFDI"))),
                    FolioUUID = tabla.GetValor(fila, "Folio UUID CFDI")?.ToString(),
                    SerieFolioInterno = tabla.GetValor(fila, "Serie Folio Interno CFDI")?.ToString(),
                    RfcEmisor = tabla.GetValor(fila, "RFC Emisor CFDI")?.ToString(),
                    NombreEmisor = tabla.GetValor(fila, "Nombre Emisor CFDI")?.ToString(),
                    RfcReceptor = tabla.GetValor(fila, "RFC Receptor CFDI")?.ToString(),
                    NombreReceptor = tabla.GetValor(fila, "Nombre Receptor CFDI")?.ToString(),
                    TipoCFDI = tabla.GetValor(fila, "Tipo CFDI")?.ToString(),
                    Version = tabla.GetValor(fila, "Version CFDI")?.ToString(),
                    Moneda = tabla.GetValor(fila, "Moneda CFDI")?.ToString(),
                    TipoCambio = Convert.ToDecimal(tabla.GetValor(fila, "Tipo Cambio CFDI") ?? 1),
                    FormaPago = tabla.GetValor(fila, "Forma Pago CFDI")?.ToString(),
                    MetodoPago = tabla.GetValor(fila, "Metodo Pago CFDI")?.ToString(),
                    CondicionPago = tabla.GetValor(fila, "Condicion Pago CFDI")?.ToString(),
                    Fiscal = tabla.GetValor(fila, "Fiscal")?.ToString(),
                    Estatus = tabla.GetValor(fila, "Estatus")?.ToString(),
                    TotalCFDI = Convert.ToDecimal(tabla.GetValor(fila, "TOTAL CFDI") ?? 0)
                });
            }
            return tblFiltrada;
        }
    }
}
