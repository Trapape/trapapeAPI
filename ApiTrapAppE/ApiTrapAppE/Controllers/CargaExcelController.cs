using Microsoft.AspNetCore.Mvc;

using Firebase.Auth;
using Firebase.Storage;
using System.Net;
using System;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using ApiTrapAppE.Models;
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

        [HttpPost]
        [Route("api/CargaExcel")]
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

                    ResponseModel response = new ResponseModel();
                    List<DataLoadsModel> ListData = new List<DataLoadsModel>();

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

                    response.isSucces = true;
                    response.URLExcel = downloadURL;
                    response.message = "Excel cargado correctamente.";

                    ListData = procesaExcel.ProcesaExcel(Objfile.file, nombre, response);

                    response.Data = ListData.ToArray();

                    return JsonConvert.SerializeObject(response);
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

        [HttpPost]
        [Route("api/DownloadExcel")]
        public async Task<string> DownloadExcel(string URLExcel)
        {
            try 
            {
                var client = new WebClient();

                Guid IdDoucumento = Guid.NewGuid();

                string nombre = IdDoucumento + ".xlsx";
                var fullPath = Path.GetFullPath(nombre);
                client.DownloadFileTaskAsync(URLExcel, fullPath);

             

                return "Carga Exitosa";
            }
            catch (Exception ex) 
            {
                return ex.Message.ToString();
            }
        }
    }
}
