using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Linq;

namespace Teste_Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Definindo o nome do arquivo
            string nomeArquivo = "Funcionarios.xlsx";
            /*Esta linha combina o caminho da área de trabalho (obtido com Environment.GetFolderPath(Environment.SpecialFolder.Desktop)) 
             * com o nome da pasta "Funcionarios" para formar o caminho completo da pasta onde o arquivo será salvo*/
            string caminhoDiretorio = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Funcionarios");
            //Combinando o caminho para chegar na pasta com o nome do arquivo excel
            string caminhoPlanilha = Path.Combine(caminhoDiretorio, nomeArquivo);

            Console.WriteLine("Pressione qualquer tecla para iniciar");
            Console.ReadKey();
            CriarPlanilha(caminhoPlanilha);

            Console.WriteLine("Pressione qualquer tecla para acessar a planilha\n");
            Console.ReadKey();
            AbrirPlanilha(caminhoPlanilha);

            Console.ReadKey();
        }

        private static void CriarPlanilha(string caminhoPlanilha)
        {
            var Funcionarios = new[]
            {
                new {Nome = "Pessoa1", Idade = 30, Cargo = "Fundador" },
                new {Nome = "Pessoa2", Idade = 50, Cargo = "Gerente" },
                new {Nome = "Pessoa3", Idade = 22, Cargo = "Vendedor" },
                new {Nome = "Pessoa4", Idade = 45, Cargo = "Assistente" },
                new {Nome = "Pessoa5", Idade = 22, Cargo = "Entregador" },
                new {Nome = "Pessoa6", Idade = 22, Cargo = "Secretario" },
            };

            // Definindo Licença
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

            // Nome da planilha
            var workSheet = excel.Workbook.Worksheets.Add("PlanilhaFuncionarios");

            // Definindo as propriedades da planilha
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Definindo as propriedades da primeira linha 
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Definindo o cabeçalho das planilhas (base 1)
            workSheet.Cells[1, 1].Value = "Nome";
            workSheet.Cells[1, 2].Value = "Idade";
            workSheet.Cells[1, 3].Value = "Cargo";

            workSheet.Cells["A1:C1"].Style.Font.Italic = true;

            // Incluindo dados na planilha
            int indice = 2;
            foreach (var funcionario in Funcionarios)
            {
                workSheet.Cells[indice, 1].Value = funcionario.Nome;
                workSheet.Cells[indice, 2].Value = funcionario.Idade;
                workSheet.Cells[indice, 3].Value = funcionario.Cargo;
                indice++;
            }

            // Ajustando o tamanho da coluna
            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();

            // Cria a pasta teste caso ela não exista
            string directoryPath = Path.GetDirectoryName(caminhoPlanilha);
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            // Se o arquivo existir, excluir
            if (File.Exists(caminhoPlanilha))
                File.Delete(caminhoPlanilha);

            // Criar o arquivo excel no disco físico
            FileStream objFileStream = File.Create(caminhoPlanilha);
            objFileStream.Close();

            // Escrever o conteúdo para o arquivo excel
            File.WriteAllBytes(caminhoPlanilha, excel.GetAsByteArray());

            // Fechando o arquivo excel
            excel.Dispose();

            Console.WriteLine($"Planilha criada com sucesso em: {caminhoPlanilha}\n");
        }

        private static void AbrirPlanilha(string caminhoPlanilha)
        {
            // Abrindo a planilha criada
            var arquivoExcel = new ExcelPackage(new FileInfo(caminhoPlanilha));

            // Localizando a planilha a ser acessada
            ExcelWorksheet planilhaFuncionarios = arquivoExcel.Workbook.Worksheets.FirstOrDefault();

            // Obtendo o número de linhas e colunas
            int rows = planilhaFuncionarios.Dimension.Rows;
            int cols = planilhaFuncionarios.Dimension.Columns;

            // Percorrendo as linhas e colunas da planilha
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    string conteudo = planilhaFuncionarios.Cells[i, j].Value.ToString();
                    Console.WriteLine($"{conteudo}");
                }
            }
        }
    }
}
