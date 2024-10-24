using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Threading;
using MinhaPlanilha.Modelos;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MinhaPlanilha
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("SISTEMA DE PLANILHAS ");
                string caminhoDaPlanilhaOriginal = ConfigurationManager.AppSettings["caminhoDaPlanilhaOriginal"];
                string caminhoDaPlanilhaNova = ConfigurationManager.AppSettings["caminhoDaPlanilhaNova"];
                Console.WriteLine("caminhoDaPlanilhaOriginal : " + caminhoDaPlanilhaOriginal);
                Console.WriteLine("caminhoDaPlanilhaNova : " + caminhoDaPlanilhaNova);
                CriarNovoArquivo(LerTabela(caminhoDaPlanilhaOriginal), caminhoDaPlanilhaNova);
                Console.WriteLine("SALVO COM SUCESSO");
                string path = $"{caminhoDaPlanilhaNova}_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}.xlsx";
                Console.WriteLine("CAMINHO DO NOVO ARQUIV0: " + path);
                Console.WriteLine("ENTER PARA FECHAR A JANELA");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine("erro: " + e.Message);
                Console.WriteLine("erro - InnerException: " + e.InnerException); 
                throw;
            }
            
        }
        private static List<Fornecedor> LerTabela(string caminhoDaPlanilhaOriginal)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage(caminhoDaPlanilhaOriginal))
            {
                ExcelWorksheet aba = excel.Workbook.Worksheets[0];
                int contadorColuna = aba.Dimension.End.Column;
                int contadorLinhas = aba.Dimension.End.Row;
                int colunaA = 1;
                int colunaB = 2;
                int colunaC = 3;
                int colunaD = 4;
                int colunaE = 5;
                int colunaF = 6;
                int colunaG = 7;
                int colunaH = 8;
                int colunaI = 9;
                int colunaJ = 10;

                var listaDeFornecedores = new List<Fornecedor>();
                var fornecedor = new Fornecedor();
                fornecedor.detalhe = new List<Detalhe>();

                for (int row = 5; row <= aba.Dimension.End.Row; row++)
                {
                    var contemNome = aba.Cells[row, colunaA].Value?.ToString().Contains("Nome:") == true;
                    var contemFone = aba.Cells[row, colunaB].Value?.ToString().Contains("Fone:") == true;
                    var contemTotal = aba.Cells[row, colunaC].Value?.ToString().Contains("Total:") == true;
                    var contemData = aba.Cells[row, colunaA].Value?.ToString().Contains("Data") == true;
                    int indice = 1;
                    var azulDaFonte = "FF0066FF";
                    var cinzaDaFonte = "FF999999";
                    var azulDoFundo = "FF002060";
                    var transparenteIndex = -2147483648;

                    var CorDoFundo = aba.Cells[row, colunaA].Style.Fill.BackgroundColor.Rgb;
                    var CorDaFonte = aba.Cells[row, colunaA].Style.Font.Color.Rgb;

                    if (CorDoFundo != null && CorDoFundo != string.Empty)
                    {
                        if (CorDoFundo == azulDoFundo)
                        {
                            // totoal geral 
                        }
                        if (aba.Cells[row, colunaA].Style.Fill.BackgroundColor.Indexed == transparenteIndex)
                        {
                        }
                        else
                        {
                            var FundoHex = "#" + CorDoFundo.Substring(2);
                            var tets = ColorTranslator.FromHtml(CorDoFundo);
                        }
                    }
                    if (CorDaFonte != null && CorDaFonte != string.Empty)
                    {
                        if (CorDaFonte == azulDaFonte)
                        {
                            if (contemNome == true && contemFone == true && contemTotal == true)
                            {
                                if (fornecedor.Nome != null && fornecedor.Nome != string.Empty)
                                {
                                    indice++;
                                    listaDeFornecedores.Add(fornecedor);
                                    fornecedor = new Fornecedor();
                                    fornecedor.detalhe = new List<Detalhe>();
                                }
                                fornecedor.id = indice;
                                fornecedor.Nome = aba.Cells[row, colunaA].Value.ToString();
                                fornecedor.Telefone = aba.Cells[row, colunaB].Value.ToString();
                                fornecedor.Total = aba.Cells[row, colunaC].Value.ToString();
                            }
                        }
                        else if (CorDaFonte == cinzaDaFonte)
                        {
                            // caso de letra cinza 
                            // titulos nao faz nada
                        }
                        else
                        {
                            var testeDeCor = CorDaFonte;
                        }
                    }
                    else if (aba.Cells[row, colunaA].Value?.ToString() != null &&
                        aba.Cells[row, colunaA].Value?.ToString() != string.Empty)
                    {
                        fornecedor.detalhe.Add(new Detalhe()
                        {
                            Data = aba.Cells[row, colunaA].Value?.ToString(),
                            Pagamento = aba.Cells[row, colunaB].Value?.ToString(),
                            Informacao = aba.Cells[row, colunaC].Value?.ToString(),
                            PlanoDeContas = aba.Cells[row, colunaD].Value?.ToString(),
                            FormaDePagamento = aba.Cells[row, colunaE].Value?.ToString(),
                            Vencimento = aba.Cells[row, colunaF].Value?.ToString(),
                            Valor = decimal.Parse(aba.Cells[row, colunaG].Value?.ToString(), NumberStyles.Currency, new CultureInfo("pt-BR")),
                            Juros = decimal.Parse(aba.Cells[row, colunaH].Value?.ToString(), NumberStyles.Currency, new CultureInfo("pt-BR")),
                            Desconto = decimal.Parse(aba.Cells[row, colunaI].Value?.ToString(), NumberStyles.Currency, new CultureInfo("pt-BR")),
                            ValorTotal = decimal.Parse(aba.Cells[row, colunaJ].Value?.ToString(), NumberStyles.Currency, new CultureInfo("pt-BR")),
                        });
                    }
                }
                return listaDeFornecedores;
            }
        }
        private static void CriarNovoArquivo(List<Fornecedor> itens, string caminhoDaPlanilhaNova)
        {
            if (itens.Count == 0) return;
            else
            {
                // Caminho onde o arquivo será salvo
                string path = $"{caminhoDaPlanilhaNova}_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}.xlsx";

                // Criação do novo arquivo Excel
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Itens");
                    int linha = 1;
                    worksheet.Cells[linha, 1].Value = "Nome";
                    worksheet.Cells[linha, 2].Value = "Telefone";
                    worksheet.Cells[linha, 3].Value = "Total";
                    worksheet.Cells[linha, 4].Value = "Data";
                    worksheet.Cells[linha, 5].Value = "Pagamento";
                    worksheet.Cells[linha, 6].Value = "Informacao";
                    worksheet.Cells[linha, 7].Value = "PlanoDeContas";
                    worksheet.Cells[linha, 8].Value = "FormaDePagamento";
                    worksheet.Cells[linha, 9].Value = "Vencimento";
                    worksheet.Cells[linha, 10].Value = "Valor";
                    worksheet.Cells[linha, 11].Value = "Juros";
                    worksheet.Cells[linha, 12].Value = "Desconto";
                    worksheet.Cells[linha, 13].Value = "ValorTotal";
                    linha++;
                    // Adiciona a lista de itens na planilha
                    for (int i = 0; i < itens.Count; i++) 
                    {
                        for (int j = 0; j < itens[i].detalhe.Count; j++) 
                        {

                            worksheet.Cells[linha, 1].Value = itens[i].Nome;
                            worksheet.Cells[linha, 2].Value = itens[i].Telefone;
                            worksheet.Cells[linha, 3].Value = itens[i].Total;
                            worksheet.Cells[linha, 4].Value = itens[i]?.detalhe[j]?.Data;
                            worksheet.Cells[linha, 5].Value = itens[i]?.detalhe[j]?.Pagamento;
                            worksheet.Cells[linha, 6].Value = itens[i]?.detalhe[j]?.Informacao;
                            worksheet.Cells[linha, 7].Value = itens[i]?.detalhe[j]?.PlanoDeContas;
                            worksheet.Cells[linha, 8].Value = itens[i]?.detalhe[j]?.FormaDePagamento;
                            worksheet.Cells[linha, 9].Value = itens[i]?.detalhe[j]?.Vencimento;
                            worksheet.Cells[linha, 10].Value = itens[i]?.detalhe[j]?.Valor;
                            worksheet.Cells[linha, 11].Value = itens[i]?.detalhe[j]?.Juros;
                            worksheet.Cells[linha, 12].Value = itens[i]?.detalhe[j]?.Desconto;
                            worksheet.Cells[linha, 13].Value = itens[i]?.detalhe[j]?.ValorTotal;
                            linha++;
                        }
                    }
                    // Salva o arquivo no caminho especificado
                    FileInfo file = new FileInfo(path);
                    package.SaveAs(file);
                }

                Console.WriteLine("Arquivo Excel criado e salvo em " + path);
            }
        }
    }
}
