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
        }
        public static SqliteConnection ObtenerConexion()
        {
            var conexion = new SqliteConnection(connectionString);
            conexion.Open();
            return conexion;
        }
    }
}
