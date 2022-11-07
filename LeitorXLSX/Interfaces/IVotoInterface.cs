using LeitorXLSX.Models;

namespace LeitorXLSX.Interfaces
{
    internal interface IVotoInterface
    {
        Task<List<Voto>>? GetVotosSegundoTurno();
    }
}
