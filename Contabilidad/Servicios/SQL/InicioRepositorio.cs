using Contabilidad.Modelos;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

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
        public void Eliminar(int id)
        {
            using(var coneccion = ConexionDb.ObtenerConexion())
            {
                using(var comando = coneccion.CreateCommand())
                {
                    comando.CommandText = "DELETE FROM Contribuyentes WHERE Id = @id;";
                    comando.Parameters.AddWithValue("@id", id);
                    comando.ExecuteNonQuery();
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
        public Usuario UsuarioActivo()
        {
            Usuario login = null;

            using (var coneccion = ConexionDb.ObtenerConexion())
            {
                using (var comando = coneccion.CreateCommand())
                {
                    comando.CommandText = "SELECT Nombre, Puesto, ColorPrincipal, ColorSecundario, ColorAcento, ColorElevado FROM Usuarios LIMIT 1;";
                    using(var lector = comando.ExecuteReader())
                    {
                        if (lector.Read())
                        {
                            login = new Usuario
                            {
                                Nombre = lector.GetString(0),
                                Puesto = lector.GetString(1),
                                ColorPrincipal = lector.GetString(2),
                                ColorSecundario = lector.GetString(3),
                                ColorAcento = lector.GetString(4),
                                ColorElevado = lector.GetString(5),
                            };
                        }
                    }
                }
            }

            return login;
        }
    }
}
