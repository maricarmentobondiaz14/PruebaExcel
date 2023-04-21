namespace CrudPruebas.Models
{
    public class CbcReglaInconsistencia
    {
        public int ReglaId { get; set; }

        public string TipoRegla { get; set; } = null!;

        public string Regla { get; set; } = null!;
    }
}
