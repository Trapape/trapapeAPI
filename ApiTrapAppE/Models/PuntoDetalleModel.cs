namespace ApiTrapAppE.Models
{
    public class PuntoDetalleModel
    {
        public String record_id { get; set; }
        public String address { get; set; }
        public String administrative_area { get; set; }
        public String country { get; set; }
        public String fecha { get; set; }
        public String hora { get; set; }
        public String locality { get; set; }
        public LocationModel location { get; set; }
        public String postal_code { get; set; }
        public String sublocality { get; set; }
    }
}