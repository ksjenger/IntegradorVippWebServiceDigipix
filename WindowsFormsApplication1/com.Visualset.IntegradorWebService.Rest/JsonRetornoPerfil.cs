using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegradorWebService.Rest
{
    public class Rootobject
    {
        public Datum[] Data { get; set; }
        public int Status { get; set; }
    }

    public class Datum
    {
        public string IdPerfil { get; set; }
        public string NomePerfil { get; set; }
    }
}
