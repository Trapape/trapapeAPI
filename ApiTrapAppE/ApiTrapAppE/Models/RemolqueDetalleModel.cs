namespace ApiTrapAppE.Models
{
    public class RemolqueDetalleModel
    {
        public String record_id { get; set; }
        public String alto { get; set; }
        public String ancho { get; set; }
        public String contenedorTamano { get; set; }
        public String contenedorTipo { get; set; }
        public String embalaje { get; set; }
        public String largo { get; set; }
        public Decimal peso { get; set; }
        public String piezas { get; set; }
        public Decimal volumen { get; set; }
    }
}
