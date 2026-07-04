using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contabilidad.Modelos
{
    public class ExcelDatosTabla
    {
        public string NombreTabla { get; set; }
        public List<string> Encabezados { get; set; }
        public object[,] Datos { get; set; }  // datos crudos del DataBodyRange
        public int NumFilas => Datos?.GetLength(0) ?? 0;
        public int NumColumnas => Datos?.GetLength(1) ?? 0;

        // Acceso por nombre de columna, no por índice hardcodeado
        public object GetValor(int fila, string nombreColumna)
        {
            int col = Encabezados.IndexOf(nombreColumna);
            if (col == -1)
            {
                throw new ArgumentException($"Columna '{nombreColumna}' no encontrada.");
            }
            return Datos[fila, col + 1]; // +1 porque el array es base 1
        }
    }
}
