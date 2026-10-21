using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contabilidad.Servicios.SQL
{
    internal class EmpresaDb
    {
        private static string connectionString;
        public static void Iniciar()
        {
            Excel_Function excel = new Excel_Function();

            string carpetaDatos = excel.ObtenerPropiedad("RutaDeTrabajo");
            string nombreEmpresa = excel.ObtenerPropiedad("Empresa");

            string rutaBaseDatos = Path.Combine(carpetaDatos, $"BaseDeDatos_{nombreEmpresa}.db");
            connectionString = $"Data Source={rutaBaseDatos}";

            SQLitePCL.Batteries.Init();

            CrearTablas();
        }
        public static SqliteConnection ObtenerConexion()
        {
            var conexion = new SqliteConnection(connectionString);
            conexion.Open();
            return conexion;
        }
        private static void CrearTablas()
        {
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText =
                @"
                CREATE TABLE IF NOT EXISTS Facturas (
                    Id_Facturas       INTEGER PRIMARY KEY AUTOINCREMENT
                                              NOT NULL,
                    Clasificacion     TEXT    NOT NULL,
                    Fiscal            TEXT,
                    Estado            TEXT    NOT NULL,
                    AplicaRegistro    TEXT,
                    SerieFolioInterno TEXT,
                    RfcEmisor         TEXT    NOT NULL,
                    NombreEmisor      TEXT    NOT NULL,
                    RfcReceptor       TEXT    NOT NULL,
                    NombreReceptor    TEXT    NOT NULL,
                    Fecha             TEXT    NOT NULL,
                    Tipo              TEXT    NOT NULL,
                    Uso               TEXT    NOT NULL,
                    Version           TEXT    NOT NULL,
                    FolioUuid         TEXT    NOT NULL
                                              UNIQUE,
                    Moneda            TEXT    NOT NULL,
                    TipoCambio        NUMERIC NOT NULL,
                    FormaPago         TEXT    NOT NULL,
                    MetodoPago        TEXT    NOT NULL,
                    CondicionPago     TEXT    NOT NULL,
                    SubTotal          NUMERIC,
                    Descuento         NUMERIC,
                    Total             NUMERIC,
                    LugarExpedicion   TEXT,
                    Exportacion       INTEGER,
                    FechaTimbrado     TEXT    NOT NULL,
                    PorcionPago       NUMERIC
                );
                
                CREATE TABLE IF NOT EXISTS ConceptosFacturas (
                    Id_ConceptosFacturas INTEGER PRIMARY KEY AUTOINCREMENT
                                                 NOT NULL,
                    Id_Factura           INTEGER REFERENCES Facturas (Id_Facturas),
                    ClaveProducto        TEXT    NOT NULL,
                    Clasificacion        TEXT,
                    Descripcion          TEXT    NOT NULL,
                    Cantidad             NUMERIC,
                    ClaveUnidad          TEXT,
                    Unidad               TEXT,
                    ValorUnitario        NUMERIC,
                    Importe              NUMERIC,
                    ObjetoImpuesto       TEXT    NOT NULL,
                    BaseImpuesto         NUMERIC,
                    CodigoImpuesto       TEXT,
                    TipoFactor           TEXT,
                    TasaoCuota           NUMERIC,
                    ImporteImpuesto      NUMERIC
                );
            
                CREATE TABLE IF NOT EXISTS ComplementosPago (
                    Id_FacturasRelacionadas INTEGER PRIMARY KEY AUTOINCREMENT,
                    Id_Facturas             INTEGER REFERENCES Facturas (Id_Facturas),
                    FechaDelPago            TEXT    NOT NULL,
                    FormaDePago             TEXT    NOT NULL,
                    Monto                   NUMERIC,
                    UuidRelacionado         TEXT,
                    SaldoAnterior           NUMERIC,
                    ImportePagado           NUMERIC,
                    SaldoInsoluto           NUMERIC,
                    Id_FacturaRelacionada   INTEGER REFERENCES Facturas (Id_Facturas) 
                );

                CREATE TABLE IF NOT EXISTS MovimientosBancarios (
                    Id_MovimientosBancarios INTEGER PRIMARY KEY AUTOINCREMENT,
                    FechaOperacion          TEXT,
                    Concepto                TEXT,
                    Referencia              TEXT,
                    Retiro                  NUMERIC,
                    Deposito                NUMERIC,
                    SaldoOperacion          NUMERIC,
                    SaldoLiquidacion        NUMERIC,
                    Año                     INTEGER,
                    Mes                     INTEGER,
                    Banco                   TEXT,
                    NumeroCuenta            TEXT,
                    Clasificacion           TEXT
                );

                CREATE TABLE IF NOT EXISTS ConciliacionBancos (
                    Id_ConciliacionBancos INTEGER PRIMARY KEY AUTOINCREMENT,
                    Id_Factura            INTEGER REFERENCES Facturas (Id_Facturas),
                    Id_MovBancario        INTEGER REFERENCES MovimientosBancarios (Id_MovimientosBancarios) 
                );";

                command.ExecuteNonQuery();
            }
        }
    }
}
