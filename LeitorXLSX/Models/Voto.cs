using System.ComponentModel.DataAnnotations;

namespace LeitorXLSX.Models
{
    public class Voto
    {
        [Key]
        public int? VotosId { get; set; }
        public int Turno { get; set; } = 0;
        public string? NomeMunicipio { get; set; } = null;
        public string? QtdAptosMunicipio { get; set; } = null;
        public string? CodigoMunicipioIBGE { get; set; } = null;
        public bool IsCapital { get; set; } = false;
        public int? Zona { get; set; } = 0;
        public int? Secao { get; set; } = 0;
        public int? QtdAptos { get; set; } = 0;
        public int? QtdVotos13 { get; set; } = 0;
        public int? QtdVotos22 { get; set; } = 0;
        public int? QtdTotalVotos1322 { get; set; } = 0;
        public int? QtdVotosBranco { get; set; } = 0;
        public int? QtdTotalFinal { get; set; } = 0;
    }
}
