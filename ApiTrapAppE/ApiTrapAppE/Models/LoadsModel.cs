namespace ApiTrapAppE.Models
{
    public class LoadsModel
    {

        public String IdLoad { get; set; }
        public Boolean Internacional { get; set; }
        public PuntoModel Punto { get; set; }
        public String actInternacional { get; set; }
        public Boolean cargaPeligrosa { get; set; }
        public Boolean cargaRefrigerada { get; set; }
        public String cargaTitulo { get; set; }
        public ConfigModel config { get; set; }
        public string dimencionAlto { get; set; }
        public string dimencionAncho { get; set; }
        public string dimensionLargo { get; set; }
        public Decimal distanciaKM { get; set; }
        public Boolean fotos { get; set; }
        public Boolean permisosEspeciales { get; set; }
        public Decimal pesoTotal { get; set; }
        public Decimal precioViaje { get; set; }
        public Boolean recibirOfertas { get; set; }
        public Boolean recomenFragil { get; set; }
        public Boolean recomenManejoCuidado { get; set; }
        public Boolean recomenMantenerSeco { get; set; }
        public String tiempoRuta { get; set; }
        public String userConsig { get; set; }
        public String userOperador { get; set; }
        public String userTranspor { get; set; }
        public String userTruck { get; set; }
        public String numRemolques { get; set; }
        public String tipoCarga { get; set; }

    }
}
