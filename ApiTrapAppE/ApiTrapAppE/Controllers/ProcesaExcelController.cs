using ApiTrapAppE.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using FireSharp;
using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;

namespace ApiTrapAppE.Controllers
{
    public class ProcesaExcelController : Controller
    {
        IFirebaseClient cliente;

        public IActionResult Index()
        {
            return View();
        }

        public string ProcesaExcel([FromForm] IFormFile ArchivoExcel, string NombreArchivo)
        {
            Stream stream = ArchivoExcel.OpenReadStream();

            IWorkbook MiExcel = null;

            if (Path.GetExtension(NombreArchivo) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MiExcel = new HSSFWorkbook(stream);
            }

            ISheet HojaExcel = MiExcel.GetSheetAt(0); 
            int cantidadFilas = HojaExcel.LastRowNum;
            int filaLoad = 0;

            List<LoadsModel> Load = new List<LoadsModel>();

            for (int i = 1; i <= cantidadFilas; i++)
            {
                IRow fila = HojaExcel.GetRow(i);

                Load.Add(new LoadsModel
                {   
                        Internacional = Convert.ToBoolean(fila.GetCell(1).ToString())
                    ,   Punto = GetPunto(fila)
                    ,   actInternacional = fila.GetCell(32).ToString()
                    ,   cargaPeligrosa = Convert.ToBoolean(fila.GetCell(33).ToString())
                    ,   cargaRefrigerada = Convert.ToBoolean(fila.GetCell(34).ToString())
                    ,   cargaTitulo = fila.GetCell(35).ToString()
                    ,   config = GetConfig(fila)
                    ,   dimencionAlto = fila.GetCell(39).ToString()
                    ,   dimencionAncho = fila.GetCell(40).ToString()
                    ,   dimensionLargo = fila.GetCell(41).ToString()
                    ,   distanciaKM = Convert.ToDecimal(fila.GetCell(42).ToString())
                    ,   fotos = Convert.ToBoolean(fila.GetCell(43).ToString())
                    ,   permisosEspeciales = Convert.ToBoolean(fila.GetCell(44).ToString())
                    ,   pesoTotal = Convert.ToDecimal(fila.GetCell(45).ToString())
                    ,   precioViaje = Convert.ToDecimal(fila.GetCell(46).ToString())
                    ,   recibirOfertas = Convert.ToBoolean(fila.GetCell(47).ToString())
                    ,   recomenFragil = Convert.ToBoolean(fila.GetCell(48).ToString())
                    ,   recomenManejoCuidado = Convert.ToBoolean(fila.GetCell(49).ToString())
                    ,   recomenMantenerSeco = Convert.ToBoolean(fila.GetCell(50).ToString())
                    ,   tiempoRuta = fila.GetCell(51).ToString()
                    ,   userConsig = fila.GetCell(52).ToString()
                    ,   userOperador = fila.GetCell(53).ToString()
                    ,   userTranspor = fila.GetCell(54).ToString()
                    ,   userTruck = fila.GetCell(55).ToString()

                });

                Object Loadfila = Load[0];

                SubirInfo(Loadfila);
            }

            return JsonConvert.SerializeObject(Load);
        }

        public PuntoModel GetPunto(IRow dr)
        {
            List<PuntoModel> Punto = new List<PuntoModel>();
            Punto.Add(new PuntoModel
            {
                    entrega = (EntregaModel)GetEntrega(dr)
                ,   recoleccion = (RecoleccionModel)GetRecoleccion(dr)

            });

            return Punto[0];
        }

        public ConfigModel GetConfig(IRow dr)
        {
            List<ConfigModel> Config = new List<ConfigModel>();
            Config.Add(new ConfigModel
            {
                config = (ConfigConfigModel)GetConfigConfig(dr)
            });

            return Config[0];
        }

        public Object GetEntrega(IRow dr)
        {
            object PEntrega = (new EntregaModel
            {
                    administrative_area = Convert.ToString(dr.GetCell(2).ToString())
                ,   country = Convert.ToString(dr.GetCell(3).ToString())
                ,   direccion = dr.GetCell(4).ToString()
                ,   fechaInicial = dr.GetCell(5).ToString()
                ,   formaCarga = dr.GetCell(6).ToString()
                ,   horaInicial = dr.GetCell(7).ToString()
                ,   latitud = Convert.ToDecimal(dr.GetCell(8).ToString())
                ,   locality = dr.GetCell(9).ToString()
                ,   longitud = Convert.ToDecimal(dr.GetCell(10).ToString())
                ,   lugarEntrega = dr.GetCell(11).ToString()
                ,   postal_code= dr.GetCell(12).ToString()
                ,   route = dr.GetCell(13).ToString()
                ,   street_number = dr.GetCell(14).ToString()
                ,   sublocality = dr.GetCell(15).ToString()
                ,   tiempoCarga = dr.GetCell(16).ToString()

            });
            return PEntrega;
        }

        public Object GetRecoleccion(IRow dr)
        {
            object PRecoleccion = (new RecoleccionModel
            {
                    administrative_area = Convert.ToString(dr.GetCell(17).ToString())
                ,   country = Convert.ToString(dr.GetCell(18).ToString())
                ,   direccion = dr.GetCell(19).ToString()
                ,   fechaInicial = dr.GetCell(20).ToString()
                ,   formaCarga = dr.GetCell(21).ToString()
                ,   horaInicial = dr.GetCell(22).ToString()
                ,   latitud = Convert.ToDecimal(dr.GetCell(23).ToString())
                ,   locality = dr.GetCell(24).ToString()
                ,   longitud = Convert.ToDecimal(dr.GetCell(25).ToString())
                ,   lugarEntrega = dr.GetCell(26).ToString()
                ,   postal_code = dr.GetCell(27).ToString()
                ,   route = dr.GetCell(28).ToString()
                ,   street_number = dr.GetCell(29).ToString()
                ,   sublocality = dr.GetCell(30).ToString()
                ,   tiempoCarga = dr.GetCell(31).ToString()

            });
            return PRecoleccion;
        }

        public ProcesaExcelController()
        {
            IFirebaseConfig config = new FirebaseConfig
            {
                AuthSecret = "ROXWHVG92cDBzvSNLDp76a4WMXgQdW36NoWnxKVi",
                BasePath = "https://trapape-default-rtdb.firebaseio.com/"

            };

            cliente = new FirebaseClient(config);
        }

        public Object GetConfigConfig(IRow dr)
        {
            object CConfig = (new ConfigConfigModel
            {
                    estatusCarga = dr.GetCell(36).ToString()
                ,   fechaCreado = Convert.ToString(dr.GetCell(37).ToString())
                ,   notificacionOferta = Convert.ToBoolean(dr.GetCell(38).ToString())
            });
            return CConfig;
        }

        public string SubirInfo(Object Load)
        {
            string IdGenerado = Guid.NewGuid().ToString("N");

            SetResponse response = cliente.Set("projects/proj_meqjHnqVDFjzhizHdj6Fjq/data/Loads/" + IdGenerado, Load);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return "Carga Exitosa";
            }
            else
            {
                return "No se cargaron los datos.";
            }
        }
    }
}
