using System.ComponentModel.DataAnnotations;

namespace LeitorXLSX.Models
{
    public class Voto
    {
        [Key]
        public int? VotosId { get; set; }
        public int Turno { get; set; } = 0;
        public string? UF { get; set; } = null;
        public string? NomeMunicipio { get; set; } = null;
        public string? QtdAptosMunicipio { get; set; } = null;
        public double? CodigoMunicipioIBGE { get; set; } = 0;
        public bool IsCapital { get; set; } = false;
        public double? Zona { get; set; } = 0;
        public double? Secao { get; set; } = 0;
        public double? QtdAptos { get; set; } = 0;
        public double? QtdVotos13 { get; set; } = 0;
        public double? QtdVotos22 { get; set; } = 0;
        public double? QtdTotalVotos1322 { get; set; } = 0;
        public double? QtdVotosBranco { get; set; } = 0;
        public double? QtdTotalFinal { get; set; } = 0;
    }
}
