using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Text.RegularExpressions;

using Firebase.Auth;
using Firebase.Storage;
using static System.Net.Mime.MediaTypeNames;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using ApiTrapAppE.Models;

namespace ApiTrapAppE.Controllers
{
    [Route("api/[controller]")]
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

        [HttpPost]
        public async Task<string> Post([FromForm] excel Objfile) 
        {
            try
            {
                if(Objfile.file.Length > 0)
                {
                    //DATOS DE CONEXION
                    string user = "cargaeexcel@gmail.com";
                    string pass = "tr4p4p3AP1#";
                    string ruta = "trapape.appspot.com";
                    string api_key = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

                    Stream archivo = Objfile.file.OpenReadStream();

                    string ext = Path.GetExtension(Objfile.file.FileName);
                    Guid IdDoucumento = Guid.NewGuid();

                    string nombre = IdDoucumento + ext;
                    
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
                        .Child("media/proj_meqjHnqVDFjzhizHdj6Fjq/app_1pAvW9AC5LiQYhzw2dpdJw/dataApplications")
                        .Child(nombre)
                        .PutAsync(archivo, cancellation.Token);

                    var downloadURL = await task;

                    var procesaExcel = new ProcesaExcelController();
                    
                    string result = procesaExcel.ProcesaExcel(Objfile.file, nombre);

                    return downloadURL;
                }
                else
                {
                    return "No se cargo el archivo correctamente.";
                }
            }catch (Exception ex)
            {
                return ex.Message.ToString();
            }
        }
    }
}
