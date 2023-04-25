namespace ApiTrapAppE.Models
{
    public class LoadsModel
    {
        public String IdLoad { get; set; }
        public string cargaDescripcion { get; set; }
        public Boolean cargaRefrigerada { get; set; }
        public string cargaTitulo { get; set; }
        public Decimal distanciaKM { get; set; }
        public Boolean foto1 { get; set; }
        public Boolean foto2 { get; set; }
        public Boolean foto3 { get; set; }
        public Boolean fotos { get; set; }
        public int numRemolques { get; set; }
        public Decimal precioViaje { get; set; }
        public Boolean recibirOfertas { get; set; }
        public Boolean recomenEstibar { get; set; }
        public Boolean recomenFragil { get; set; }
        public Boolean recomenManejoCuidado { get; set; }
        public Boolean recomenMantenerSeco { get; set; }
        public String seguroCarga { get; set; }
        public String tiempoRuta { get; set; }
        public String tipoCarga { get; set; }
        public String tipoUnidad { get; set; }
        public String userConsig { get; set; }
        public String userOperador { get; set; }
        public String userTranspor { get; set; }
        public PuntoModel Punto { get; set; }
        public RemolqueModel Remolque { get; set; }
        public ConfigModel config { get; set; }
        public string nombreExcel { get; set; }
    }
}
