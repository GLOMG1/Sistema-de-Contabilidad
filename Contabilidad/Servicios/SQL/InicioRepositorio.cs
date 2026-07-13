using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Contabilidad.Servicios.SQL
{
    internal class InicioRepositorio
    {
        public void Insertar(string Nombre, string Rfc, int Regimen)
        {
            try
            {
                using (var connection = ConexionDb.ObtenerConexion())
                {

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
        public List<Contribuyentes> Consultar()
        {
            var lista = new List<Contribuyentes>();
            using (var connection = ConexionDb.ObtenerConexion())
            {

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
