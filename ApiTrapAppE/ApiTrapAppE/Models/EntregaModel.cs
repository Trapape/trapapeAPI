﻿namespace ApiTrapAppE.Models
{
    public class EntregaModel
    {
        public String record_id { get; set; }
        public String address { get; set; }
        public String administrative_area { get; set; }
        public String country { get; set; }
        public String fecha { get; set; }
        public String hora { get; set; }
        public String locality { get; set; }
        public Decimal latitud { get; set; }
        public Decimal longitud { get; set; }
        public String postal_code { get; set; }
        public String sublocality { get; set; }
    }
}
