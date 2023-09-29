namespace ApiTrapAppE.Models
{
    public class ConfigConfigModel
    {
        public string record_id { get; set; }
        public string estatusCarga { get; set; }
        public string estatus { get; set; }
        public string fechaActualizacion { get; set; }
        public string fechaCreado { get; set; }
        public bool notificacionOferta { get; set; }
        public bool privacidad { get; set; }

        // Constructor que asigna valores predeterminados
        public ConfigConfigModel()
        {
            record_id = "config";
            estatusCarga = "Publicada";
            estatus = "Publicada";
            notificacionOferta = false;
            privacidad = false;
        }
    }
}