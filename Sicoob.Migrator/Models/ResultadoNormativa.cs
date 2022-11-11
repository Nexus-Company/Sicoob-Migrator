using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sicoob.Migrator.Models
{
    internal class ResultadoNormativo
    {
        public Normativo Normativo { get; set; }
        public DateTime Date { get; } = DateTime.Now;
        public bool Success => string.IsNullOrEmpty(Error);
        public string Error { get; set; }
        public ResultadoNormativo(Normativo normativo)
            : this(normativo, string.Empty)
        { }
        public ResultadoNormativo(Normativo normativo, string error)
        {
            Normativo = normativo ?? throw new ArgumentNullException(nameof(normativo));
            Error = error ?? throw new ArgumentNullException(nameof(error));
        }
    }
}
