using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using System.Windows;

namespace Contabilidad.Servicios
{
    public class Contribuyentes
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public string Rfc { get; set; }
        public int Regimen { get; set; }
    }

    internal class DbContabilidad
    {
        private const string connectionString = "Data Source=dbContabilidad.db";

        public static void Iniciar()
        {
            SQLitePCL.Batteries.Init();

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
                    Regimen INTEGER NOT NULL
                );";

                command.ExecuteNonQuery();
            }
        }

        public static void Insertar(string Nombre, string Rfc, int Regimen)
        {
            try
            { 
                using (var connection = new SqliteConnection(connectionString))
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                    command.CommandText =
                        @"
                        INSERT INTO Contribuyentes (Nombre, Rfc, Regimen)
                        VALUES ($Nombre, $Rfc, $Regimen)";

                        command.Parameters.AddWithValue("$Nombre", Nombre);
                        command.Parameters.AddWithValue("$Rfc", Rfc);
                        command.Parameters.AddWithValue("$Regimen", Regimen);

                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (SqliteException ex)
            {
                if (ex.SqliteErrorCode == 19)
                {
                    MessageBox.Show("El RFC ya existe en la base de datos.", "Dato Duplicado.", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show("Error al insertar: " + ex.Message);
                }
            }
        }

        public static List<Contribuyentes> Consultar()
        {
            var lista = new List<Contribuyentes>();
            using (var connection = new SqliteConnection(connectionString))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT Id, Nombre, Rfc, Regimen FROM Contribuyentes;";
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            lista.Add(new Contribuyentes
                            {
                                Id = reader.GetInt32(0),
                                Nombre = reader.GetString(1),
                                Rfc = reader.GetString(2),
                                Regimen = reader.GetInt32(3)
                            });
                        }
                    }
                }
            }
            return lista;
        }
    }
}
