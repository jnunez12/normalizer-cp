using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NormalizarCP.Entidades
{
    public class Calle
    {
        public int id { get; set; }
        public string nro_zona {get;set;}
        public string calle { get; set; }
        public int altura_ini { get; set; }
        public string cp { get; set; }
    }
}
