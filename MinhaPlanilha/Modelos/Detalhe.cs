using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinhaPlanilha.Modelos
{
    public class Detalhe
    {
        public string Data { get; set; }
        public string Pagamento { get; set; }
        public string Informacao { get; set; }
        public string PlanoDeContas { get; set; }
        public string FormaDePagamento { get; set; }
        public string Vencimento { get; set; }
        public decimal Valor { get; set; }
        public decimal Juros { get; set; }
        public decimal Desconto { get; set; }
        public decimal ValorTotal { get; set; }

    }
}
