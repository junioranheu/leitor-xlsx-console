using LeitorXLSX.Data;
using LeitorXLSX.Interfaces;
using LeitorXLSX.Models;
using LeitorXLSX.Repositories;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

Console.WriteLine("Iniciando projeto");
var config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: false)
            .AddUserSecrets<Program>()
            .Build();

Console.WriteLine("Iniciando conexão com a base de dados");
string con = config.GetConnectionString("BaseDados"); // Connection string;
var secretPass = config["SecretSenhaBancoDados"]; // secrets.json;
con = con.Replace("[secretSenhaBancoDados]", secretPass); // Alterar pela senha do secrets.json;

var serviceProvider = new ServiceCollection()
                     .AddDbContext<Context>(options => options.UseMySql(con, ServerVersion.AutoDetect(con)))
                     .AddSingleton<IVotoInterface, VotoRepository>()
                     // .AddSingleton<ITesteInterface, BarService>()
                     .BuildServiceProvider();

bool isContinuar = true;
bool resetarBd = false;
if (resetarBd)
{
    try
    {
        Console.WriteLine("Restaurando a base de dados");
        var context = serviceProvider.GetRequiredService<Context>();

        await context.Database.EnsureDeletedAsync();
        await context.Database.EnsureCreatedAsync();
    }
    catch (Exception ex)
    {
        string erroBD = ex.Message.ToString();
        Console.WriteLine($"Falha ao resetar a base de dados: {erroBD}");
        isContinuar = false;
    }
}

if (isContinuar)
{
    Console.WriteLine("\nIniciando leitura dos arquivos XLSXs");
    var votos = serviceProvider.GetService<IVotoInterface>();
    IEnumerable<Voto> xlsxVotos = await votos?.GetVotosSegundoTurno();

    Console.WriteLine($"\nResultado: {xlsxVotos?.Count()}");
    //if (xlsxVotos is not null)
    //{
    //    foreach (var item in xlsxVotos)
    //    {
    //        string? nomeMunicipio = item?.NomeMunicipio == "-" ? "N/A" : $"+{item?.NomeMunicipio}";
    //        double? zona = item?.Zona == 0 ? 0 : item?.Zona;
    //        double? secao = item?.Secao == 0 ? 0 : item?.Secao;

    //        Console.WriteLine($"Cidade: {nomeMunicipio} | Zona: {zona} | Seção: {secao}");
    //    }
    //}
}

Console.WriteLine("\nFim");