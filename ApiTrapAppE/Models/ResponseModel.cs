namespace ApiTrapAppE.Models
{
    public class ResponseModel
    {
        public bool isSucces { get; set; }
        public string message { get; set; }
        public string URLExcel { get; set; }
        public DataLoadsModel[] Data { get; set; }
    }
}