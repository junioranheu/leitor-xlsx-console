using LeitorXLSX.Models;

namespace LeitorXLSX.Interfaces
{
    internal interface IVotoInterface
    {
        Task<IEnumerable<Voto>>? GetVotosSegundoTurno();
    }
}
