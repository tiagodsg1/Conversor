using System.Text;
using System.Threading.Tasks;
using System.IO;
using System;
using System.Diagnostics;
using ClosedXML.Excel;



namespace Contador
{
    class Program
    {
        static void Main(string[] args)        
        {

            
            string ano_pasta, nome_pasta; // Declarando variavel 
            string[] arquivos = Directory.GetDirectories(@"C:\Users\tiago\Desktop\ClouGed");// Entrando na pasta 
            foreach (var arquivo in arquivos)
            {

                Console.WriteLine(arquivo);//Printando os arquivos que tem dentro do caminho a cima
                

            }
            Console.WriteLine("Digite o ano desejado, (ex: 2022)");
            ano_pasta = Console.ReadLine();// Dando valor do ano da pasta para a variavel ano_pasta 
            nome_pasta = @"C:\Users\tiago\Desktop\ClouGed\" +ano_pasta;//Concatenando o caminho da pasta com o ano e atribuindo valor para a variavel nome_pasta
            Console.WriteLine(nome_pasta);             
            foreach (var arquivo in arquivos)
            {
                if (arquivo == nome_pasta)
                {

                    string[] arquivos_doc = Directory.GetFiles(nome_pasta);
                    string[] nomes_Arquivo = arquivos_doc;
                    
                    
                    for (int i = 0; i < nomes_Arquivo.Length; i++)
                    {
                        string[] allLines = File.ReadAllLines(nomes_Arquivo[i]);
                        string linha2 = allLines[1];
                        string linha4 = allLines[3];
                        string linha34 = allLines[33];
                        Console.WriteLine(linha2);
                        Console.WriteLine(linha4);
                        Console.WriteLine(linha34);
                        string periodo_referencia = linha2.Substring(23, 8);
                        Console.WriteLine(periodo_referencia);
                        string inscricao = linha4.Substring(24, 9);
                        Console.WriteLine(inscricao);
                        string razao_social = linha4.Substring(36);
                        Console.WriteLine(razao_social);
                        string saidas_tributarias = linha34.Substring(25, 21);
                        Console.WriteLine(saidas_tributarias);

                        string planilha;
                        planilha = nome_pasta + "Planilha.xlsx";                        
                        using (var workbook = new XLWorkbook())
                        {
                            int j = 0;
                            var worksheet = workbook.Worksheets.Add("Planilha");
                            worksheet.Cell("A" +(i+ 1)).Value = periodo_referencia[j];                            


                            workbook.SaveAs(planilha);
                            j++;

                        }
                        Process.Start(new ProcessStartInfo(planilha) { UseShellExecute = true });
                    }


                    

                }

                
            }

            
            Console.ReadKey();
           
        }
    }
}
