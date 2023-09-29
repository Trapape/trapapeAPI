using ApiTrapAppE.Models;
using Firebase.Auth;
using Firebase.Storage;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace ApiTrapAppE.Controllers
{
    [ApiController]
    public class CargaExcelController : ControllerBase
    {
        public class excel
        {
            public IFormFile file { get; set; }
        }

        public class nombre
        {
            public string namefile { get; set; }
        }

        public static IWebHostEnvironment _env { get; set; }

        public CargaExcelController(IWebHostEnvironment env)
        {
            _env = env;
        }

        [HttpPost]
        [Route("api/CargaExcel")]
        public async Task<string> Post([FromForm] excel Objfile, [FromForm] string userConsig)
        {
            try
            {
                if (Objfile.file.Length > 0)
                {
                    //DATOS DE CONEXION
                    string user = "cargaeexcel@gmail.com";
                    string pass = "tr4p4p3AP1#";
                    string ruta = "trapape.appspot.com";
                    string api_key = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

                    Stream archivo = Objfile.file.OpenReadStream();

                    ResponseModel response = new ResponseModel();
                    List<DataLoadsModel> ListData = new List<DataLoadsModel>();

                    string ext = Path.GetExtension(Objfile.file.FileName);
                    Guid IdDoucumento = Guid.NewGuid();

                    string nombre = IdDoucumento + ext;
                    string idCargaPrincipal = Convert.ToString(IdDoucumento);

                    var auth = new FirebaseAuthProvider(new FirebaseConfig(api_key));
                    var access = await auth.SignInWithEmailAndPasswordAsync(user, pass);

                    var cancellation = new CancellationTokenSource();

                    var task = new FirebaseStorage(
                        ruta,
                        new FirebaseStorageOptions
                        {
                            AuthTokenAsyncFactory = () => Task.FromResult(access.FirebaseToken),
                            ThrowOnCancel = true
                        })
                        .Child("media/proj_meqjHnqVDFjzhizHdj6Fjq/app_vjubyyTnE5REBNbo1HHscW/dataApplications")
                        .Child(nombre)
                        .PutAsync(archivo, cancellation.Token);

                    var downloadURL = await task;

                    var procesaExcel = new ProcesaExcelController();

                    response.isSucces = true;
                    response.URLExcel = downloadURL;
                    response.message = "Excel cargado correctamente.";

                    ListData = procesaExcel.ProcesaExcel(Objfile.file, nombre, userConsig, downloadURL, idCargaPrincipal);

                    response.Data = ListData.ToArray();

                    return JsonConvert.SerializeObject(response);
                }
                else
                {
                    return "No se cargo el archivo correctamente.";
                }
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }

        [HttpPost]
        [Route("api/DownloadExcel")]
        public async Task<IActionResult> DownloadExcel(string URLExcel, string userConsig)
        {
            try
            {
                // Validar y deshacer la codificación de la URL
                URLExcel = Uri.UnescapeDataString(URLExcel);
                URLExcel = URLExcel.Replace("%2F", "/");

                // Extraer el nombre del archivo Excel
                var _urlsplit = URLExcel.Split("/");
                var nombre = _urlsplit.LastOrDefault()?.Split("?").FirstOrDefault() ?? "";

                if (string.IsNullOrEmpty(nombre))
                {
                    // Manejar el caso en el que no se pudo extraer el nombre del archivo
                    return BadRequest("No se pudo determinar el nombre del archivo.");
                }

                // Configuración de Firebase
                string user = "cargaeexcel@gmail.com";
                string pass = "tr4p4p3AP1#";
                string ruta = "trapape.appspot.com";
                string api_key = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

                Guid IdDoucumento = Guid.NewGuid();
                string idCargaPrincipal = IdDoucumento.ToString();

                // Autenticación en Firebase
                var auth = new FirebaseAuthProvider(new FirebaseConfig(api_key));
                var access = await auth.SignInWithEmailAndPasswordAsync(user, pass);

                // Configuración de FirebaseStorage
                var storage = new FirebaseStorage(
                    ruta,
                    new FirebaseStorageOptions
                    {
                        AuthTokenAsyncFactory = () => Task.FromResult(access.FirebaseToken),
                        ThrowOnCancel = true
                    });

                // Obtener la URL de descarga
                string task = await storage
                    .Child("media/proj_meqjHnqVDFjzhizHdj6Fjq/app_vjubyyTnE5REBNbo1HHscW/dataApplications")
                    .Child(nombre)
                    .GetDownloadUrlAsync();

                // Descargar el archivo
                using (var client = new HttpClient())
                {
                    var httpResponse = await client.GetAsync(task);
                    var streamToReadFrom = await httpResponse.Content.ReadAsStreamAsync();

                    if (!nombre.EndsWith(".xlsx"))
                    {
                        nombre += ".xlsx";
                    }

                    // Crear un objeto IFormFile a partir del flujo
                    var objFile = new FormFile(streamToReadFrom, 0, streamToReadFrom.Length, null, nombre);

                    // Lógica de procesamiento en una clase separada (ProcesaExcelController)
                    var procesaExcel = new ProcesaExcelController();
                    var response = procesaExcel.ProcesaExcel(objFile, nombre, userConsig, task, idCargaPrincipal);

                    return Ok(response); // Devolver una respuesta HTTP 200 OK con el objeto response serializado automáticamente a JSON.
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message); // Devolver una respuesta HTTP 400 Bad Request con el mensaje de error.
            }
        }

        [HttpGet]
        [Route("api/GetPrueba")]
        public async Task<string> GetPrueba(string parametro_prueba)
        {
            try
            {
                return "Get exitoso, parametro: " + parametro_prueba;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
    }
}