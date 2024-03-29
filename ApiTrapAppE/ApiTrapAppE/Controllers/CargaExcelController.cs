﻿using Microsoft.AspNetCore.Mvc;

using Firebase.Auth;
using Firebase.Storage;
using System.Net;
using ApiTrapAppE.Models;
using Newtonsoft.Json;
using System.Xml.Linq;
using System.Security.Policy;
using System;

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

                    ListData = procesaExcel.ProcesaExcel(Objfile.file, nombre, response, userConsig, downloadURL, idCargaPrincipal);

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
        public async Task<string> DownloadExcel(string URLExcel, string userConsig)
        {
            try 
            {
                URLExcel = URLExcel.Replace("%2F", "/");
                var _urlsplit = URLExcel.Split("/");
                var url = _urlsplit[_urlsplit.Length - 1];
                _urlsplit = url.Split("?");

                string nombre = "";

                foreach (var item in _urlsplit)
                {
                    if (item.Contains(".xlsx"))
                    {
                        nombre = item;
                    }

                    if (nombre == "")
                    {
                        nombre = item;
                    }
                }

                string user = "cargaeexcel@gmail.com";
                string pass = "tr4p4p3AP1#";
                string ruta = "trapape.appspot.com";
                string api_key = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

                Guid IdDoucumento = Guid.NewGuid();

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
                        .GetDownloadUrlAsync().GetAwaiter().GetResult();

                Stream stream = null;
                string strURL = URLExcel;
                HttpClient client = new HttpClient();
                HttpResponseMessage httpResponse = await client.GetAsync(task);
                Stream streamToReadFrom = await httpResponse.Content.ReadAsStreamAsync();

                if (!nombre.Contains(".xlsx"))
                {
                    nombre = nombre + ".xlsx";
                }

                IFormFile Objfile = new FormFile(streamToReadFrom, 0, streamToReadFrom.Length, nombre.Replace(".xlsx", ""), nombre);

                ResponseModel response = new ResponseModel();
                List<DataLoadsModel> ListData = new List<DataLoadsModel>();

                var procesaExcel = new ProcesaExcelController();

                response.isSucces = true;
                response.URLExcel = task;
                response.message = "Documento Procesado.";

                ListData = procesaExcel.ProcesaExcel(Objfile, nombre, response, userConsig, task, idCargaPrincipal);

                response.Data = ListData.ToArray();

                return JsonConvert.SerializeObject(response);
            }
            catch (Exception ex) 
            {
                return ex.Message.ToString();
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
