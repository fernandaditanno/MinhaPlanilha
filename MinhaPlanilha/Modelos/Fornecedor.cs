﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinhaPlanilha.Modelos
{
    public class Fornecedor
    {
        public int id { get; set; }
        public string Nome { get; set; }
        public string Telefone { get; set; }
        public string Total { get; set; }
        public List<Detalhe> detalhe { get; set; }
    }
}
