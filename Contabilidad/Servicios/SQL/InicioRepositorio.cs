using Contabilidad.Modelos;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;

namespace Contabilidad.Servicios.SQL
{
    public static class InicioRepositorio
    {
        public static void Insertar(string Nombre, string Rfc, int Regimen, string Ruta)
        {
            Ruta = VerificacionCarpeta(Ruta, Nombre);
            if (Ruta == null)
            {
                MessageBox.Show("Ya existe una carpeta con el mismo nombre en esa ruta");
                return;
            }
            try
            {
                using (var connection = ConexionDb.ObtenerConexion())
                {

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText =
                            @"
                        INSERT INTO Contribuyentes (Nombre, Rfc, Regimen, RutaDeTrabajo)
                        VALUES ($Nombre, $Rfc, $Regimen, $RutaDeTrabajo)";

                        command.Parameters.AddWithValue("$Nombre", Nombre);
                        command.Parameters.AddWithValue("$Rfc", Rfc);
                        command.Parameters.AddWithValue("$Regimen", Regimen);
                        command.Parameters.AddWithValue("$RutaDeTrabajo", Ruta);

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
        public static void Eliminar(int id)
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
        public static List<Empresa> Consultar()
        {
            var lista = new List<Empresa>();
            using (var connection = ConexionDb.ObtenerConexion())
            {

                using (var command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT Id, Nombre, Rfc, Regimen, RutaDeTrabajo FROM Contribuyentes;";
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            lista.Add(new Empresa
                            {
                                Id = reader.GetInt32(0),
                                Nombre = reader.GetString(1),
                                Rfc = reader.GetString(2),
                                Regimen = reader.GetInt32(3),
                                RutaDeTrabajo = reader.GetString(4)
                            });
                        }
                    }
                }
            }
            return lista;
        }
        public static Usuario UsuarioActivo()
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
        public static Empresa EmpresaActiva()
        {
            Empresa login = null;

            using (var coneccion = ConexionDb.ObtenerConexion())
            {
                using (var comando = coneccion.CreateCommand())
                {
                    comando.CommandText = "SELECT Nombre, Rfc, Regimen, RutaDeTrabajo FROM Contribuyentes LIMIT 1;";
                    using (var lector = comando.ExecuteReader())
                    {
                        if (lector.Read())
                        {
                            login = new Empresa
                            {
                                Nombre = lector.GetString(0),
                                Rfc = lector.GetString(1),
                                Regimen = lector.GetInt32(2),
                                RutaDeTrabajo = lector.GetString(3)
                            };
                        }
                    }
                }
            }

            return login;
        }
        private static string VerificacionCarpeta(string ruta, string nombre)
        {
            string carpetaDatos = Path.Combine(ruta, nombre);

            if (Directory.Exists(carpetaDatos))
            {
                return null;
            }
            else
            {
                Directory.CreateDirectory(carpetaDatos);
                return carpetaDatos;
            }
        }
    }
}
