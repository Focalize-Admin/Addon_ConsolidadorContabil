using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_Messem
{
    internal class Erro
    {
        public string Rotina { get; set; }
        public string ID { get; set; }
        public string Banco { get; set; }
        public string Error { get; set; }

        public Erro(string rotina, string iD, string banco, string error)
        {
            Rotina = rotina;
            ID = iD;
            Banco = banco;
            Error = error;
        }
    }
}
