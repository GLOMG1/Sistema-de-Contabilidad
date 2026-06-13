using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Contabilidad.Modelos
{
    public class ClassBancos
    {
        public int IdDataBody { get; set; }
        public int IdRegistro { get; set;  }
        public DateTime FechaOperacion {  get; set; }
        public string FechaLiquidacion { get; set; }
        public string Concepto { get; set; }
        public string Referencia { get; set; }
        public decimal Importe { get; set; }
        public decimal SaldoOperacion { get; set; }
        public decimal SaldoLiquidacion { get; set; }
        public int Año { get; set; }
        public string Mes {  get; set; }
        public string Banco {  get; set; }
        public string NumCuenta { get; set;  }
        public string Clasificacion { get; set; }
        public string SerieFolioInternoCFDI { get; set; }
        public string FechaCFDI { get; set; }
        public string FolioUUIDCFDI { get; set; }
        public string NombreEmisorReceptorCFDI { get; set; }
        public string Estatus { get; set; }
        public string IdRegistroAuxiliar { get; set; }
        public string FechaAuxiliar { get; set; }
        public string FolioAuxilar { get; set; }
        public string EstatusAUXILIAR { get; set; }

    }
}
