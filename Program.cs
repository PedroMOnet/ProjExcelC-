//Esse código possui funções apenas de aprendizado e não deve ser usado para fins lucrativos,a tribuo a autoria desse código ao copilot e o EdgarHygino.

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{

    static void Main(string[] args)
    {
        while (true)
        {
            //cria uma lista para o objeto Cadastros
            List<Cadastros> ListadePessoas = new List<Cadastros>();

            int opcao = Menu();
            //ler os usuários cadastrados na planilia
            if (opcao == 1)
            {
                Console.Clear();
                //o método try catch serve para captar possiveis excessões
                try
                {
                    //essa linha tem a finalidade de localizar o caminho até a planilia do excel
                    var Path = "C:\\Users\\pedro\\Documents\\PlaniliaExcel\\01.xlsx";
                    var xls = new XLWorkbook(Path);//essa clase é usada para poder manipular a planilha do excel na qual foi atribuida
                    var planilha = xls.Worksheets.FirstOrDefault(w => w.Name == "Plan1");
                    var totalLinhas = planilha.RowsUsed().Count();

                    if (planilha == null)
                    {
                        Console.WriteLine("planilha não encontrada");
                    }
                    else{
                    //aqui criamos um loop para percorrer a planilha e extrairmos os seus valores
                    for (int i = 2; i <= totalLinhas; i++)
                    {
                        //ler os dados da planilha
                        var NomePessoa = planilha.Cell($"A{i}").Value.ToString();
                        var idade = Convert.ToInt32(planilha.Cell($"B{i}").Value.ToString());
                        Cadastros novoCadastro = new Cadastros { Nome = NomePessoa, Idade = idade};
                        ListadePessoas.Add(novoCadastro);
                        Console.WriteLine($"{NomePessoa} - {idade}");
                    }
                    while(true){
                    Console.WriteLine("aperte qualquer tecla");
                    string espaco = Console.ReadLine();
                    if(espaco != null){break;}
                    }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ocorreu um erro: {ex.Message}");

                }
            }
            //digitar os usuários manualmente
            else if (opcao == 2)
            {
                Console.WriteLine("Quantas pessoas a serem cadastradas?");
                int totalPessoas = Convert.ToInt32(Console.ReadLine());

                for (int j = 0; j < totalPessoas; j++)
                {
                    Console.WriteLine($"Digite o nome da pessoa {j + 1}:");
                    string nomePessoa = Console.ReadLine();

                    Console.WriteLine($"digite a idade da pessoa {j + 1}:");
                    int idade = Convert.ToInt32(Console.ReadLine());

                    Cadastros novoCadastro = new Cadastros { Nome = nomePessoa, Idade = idade };
                    ListadePessoas.Add(novoCadastro);
                }
                InsercaoPessoasPlanilha(ListadePessoas);
            }

            else if (opcao == 3)
            {
                System.Environment.Exit(0);
            }


            else
            {
                Console.WriteLine("opção inválida");
            }
            Console.WriteLine("Cadastros realizados");
            foreach (var pessoa in ListadePessoas)
            {
                Console.WriteLine($"{pessoa.Nome} - {pessoa.Idade}");
            }
        }

        //função para salvsar os dados na planilha
        static void InsercaoPessoasPlanilha(List<Cadastros> ListadePessoas)
        {
            try
            {
                var path = "C:\\Users\\pedro\\Documents\\PlaniliaExcel\\01.xlsx";
                var xls = new XLWorkbook(path);
                var planilha = xls.Worksheets.First(w => w.Name == "Plan1");

                int linhaInicial = planilha.RowsUsed().Count() + 1;

                foreach (var pessoa in ListadePessoas)
                {
                    planilha.Cell($"A{linhaInicial}").Value = pessoa.Nome;
                    planilha.Cell($"B{linhaInicial}").Value = pessoa.Idade;
                    linhaInicial++;
                }

                xls.Save();
                Console.WriteLine("Dados salvos na planilha com sucesso.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocorreu um erro ao salvar os dados na planilha: {ex.Message}");
            }
        }
        static int Menu()
        {
            //inicia o programa limpando todas as mensagens do console
            Console.Clear();
            //Capita a decisão do usuário, toda vez que se terminar uma ação ele deve retornar para essa função
            Console.WriteLine("Escolha uma opção: ");
            Console.WriteLine("\n1-Ler usuários cadastrados: ");
            Console.WriteLine("\n2-Inserir os dados manualmente");
            Console.WriteLine("\n3-Sair da aplicação");
            int opcao = Convert.ToInt32(Console.ReadLine());

            return opcao;
        }
    }

    class Cadastros
    {
        public string Nome { get; set; }
        public int Idade { get; set; }
    }
}

//System.Enviroment.Exit(); <-- se usa para finalizar uma aplicação 