using Contabilidad.Modelos;
using Microsoft.Data.Sqlite;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Contabilidad.Servicios.SQL
{
    internal class ConexionDb
    {
        private static string connectionString;
        public static void Iniciar()
        {
            string carpetaDatos = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Contabilidad");

            if (!Directory.Exists(carpetaDatos))
                Directory.CreateDirectory(carpetaDatos);

            string rutaBaseDatos = Path.Combine(carpetaDatos, "dbContabilidad.db");
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
                CREATE TABLE IF NOT EXISTS Contribuyentes (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nombre TEXT NOT NULL,
                    Rfc TEXT NOT NULL UNIQUE,
                    Regimen INTEGER NOT NULL,
                    RutaDeTrabajo TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS Usuarios (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Nombre TEXT NOT NULL,
                    Puesto TEXT NOT NULL,
                    ColorPrincipal TEXT,
                    ColorSecundario TEXT,
                    ColorAcento TEXT,
                    ColorElevado TEXT
                );

                CREATE TABLE IF NOT EXISTS Filtros (
                    Nombre TEXT NOT NULL
                );
                
                INSERT OR IGNORE INTO Usuarios (
                    Nombre,
                    Puesto,
                    ColorPrincipal,
                    ColorSecundario,
                    ColorAcento,
                    ColorElevado
                )
                VALUES (
                    'Usuario predeterminado',
                    'Administrador',
                    '#FFFFFF',
                    '#000000',
                    '#007ACC',
                    '#EEEEEE'
                );";

                command.ExecuteNonQuery();
            }
        }
    }
}
