namespace ApiTrapAppE.Models
{
    public class DataLoadsModel
    {
        public bool isSucces { get; set; }
        public string message { get; set; }
        public string idLoad { get; set; }
        public LoadsModel Load { get; set; }
    }
}