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
using System.Data;

namespace ApiTrapAppE.Controllers
{
    public class ProcesaExcelController : Controller
    {
        IFirebaseClient cliente;

        public IActionResult Index()
        {
            return View();
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

            int cantidadHojas = MiExcel.NumberOfSheets; // Saca el numero de hojas para los diferentes tipos de carga

            List<LoadsModel> Load = new List<LoadsModel>();

            //Recorre las hojas para insertar de cada carga
            for (int hoja = 0; hoja < cantidadHojas; hoja++)
            {
                ISheet HojaExcel = MiExcel.GetSheetAt(hoja);
                int cantidadFilas = HojaExcel.LastRowNum;

                if(cantidadFilas == 0)
                {
                    continue;
                }

                int cuerpoFilas = 0;
                DataTable dtExcelData = GeneraTabla(HojaExcel.GetRow(0), HojaExcel.GetRow(0).LastCellNum);

                for (int ifila = 0; ifila <= cantidadFilas; ifila++)
                {
                    IRow fila = HojaExcel.GetRow(ifila);

                    if (ifila == 0)
                    {
                        continue;
                    }

                    InsertaTabla(dtExcelData, fila, fila.LastCellNum);
                }

                GeneraLoad(dtExcelData, NombreArchivo);
            }

            return JsonConvert.SerializeObject(Load);
        }

        public DataTable GeneraTabla(IRow dr, int intdtcolumn)
        {
            var dtExcelData = new DataTable();

            for (int icolumn = 0; icolumn < intdtcolumn; icolumn++)
            {
                dtExcelData.Columns.Add(dr.GetCell(icolumn).ToString());

            }

            return dtExcelData;
        }

        public DataTable InsertaTabla(DataTable dtExcelData, IRow dr, int intdtcolumn)
        {
            DataRow renglon = dtExcelData.NewRow();

            for (int icolumn = 0; icolumn < intdtcolumn; icolumn++)
            {
                if(dr.GetCell(icolumn) == null)
                {
                    renglon[icolumn] = "";
                }
                else
                {
                    renglon[icolumn] = dr.GetCell(icolumn).ToString();
                }
            }

            dtExcelData.Rows.Add(renglon);

            return dtExcelData;
        }

        public string GeneraLoad(DataTable dtExcelData, string NombreArchivo)
        {
            if(dtExcelData.Rows.Count > 0)
            {
                foreach (DataRow row in dtExcelData.Rows)
                {
                    string cargaDescripcion;
                    List<LoadsModel> Load = new List<LoadsModel>();

                    Guid IdGenerado = Guid.NewGuid();

                    Load.Add(new LoadsModel
                    {
                            IdLoad = Convert.ToString(IdGenerado)
                        ,   cargaDescripcion = Convert.ToString(row.Field<string>("cargaDescripcion"))
                        ,   cargaRefrigerada = Convert.ToBoolean(row.Field<string>("cargaRefrigerada"))
                        ,   cargaTitulo = Convert.ToString(row.Field<string>("cargaTitulo"))
                        ,   distanciaKM = Convert.ToDecimal(row.Field<string>("distanciaKM"))
                        ,   foto1 = Convert.ToBoolean(row.Field<string>("foto1"))
                        ,   foto2 = Convert.ToBoolean(row.Field<string>("foto2"))
                        ,   foto3 = Convert.ToBoolean(row.Field<string>("foto3"))
                        ,   fotos = Convert.ToBoolean(row.Field<string>("fotos"))
                        ,   numRemolques = Convert.ToInt32(row.Field<string>("numRemolques"))
                        ,   precioViaje = Convert.ToDecimal(row.Field<string>("precioViaje"))
                        ,   recibirOfertas = Convert.ToBoolean(row.Field<string>("recibirOfertas"))
                        ,   recomenEstibar = Convert.ToBoolean(row.Field<string>("recomenEstibar"))
                        ,   recomenFragil = Convert.ToBoolean(row.Field<string>("recomenFragil"))
                        ,   recomenManejoCuidado = Convert.ToBoolean(row.Field<string>("recomenManejoCuidado"))
                        ,   recomenMantenerSeco = Convert.ToBoolean(row.Field<string>("recomenMantenerSeco"))
                        ,   seguroCarga = Convert.ToString(row.Field<string>("seguroCarga"))
                        ,   tiempoRuta = Convert.ToString(row.Field<string>("tiempoRuta"))
                        ,   tipoCarga = Convert.ToString(row.Field<string>("tipoCarga"))
                        ,   tipoUnidad = Convert.ToString(row.Field<string>("tipoUnidad"))
                        ,   userConsig = Convert.ToString(row.Field<string>("userConsig"))
                        ,   userOperador = Convert.ToString(row.Field<string>("userOperador"))
                        ,   userTranspor = Convert.ToString(row.Field<string>("userTranspor"))
                        ,   Punto = GetPunto(row)
                        ,   Remolque = GetRemolque(row)
                        ,   config = GetConfig(row)
                        ,   nombreExcel = NombreArchivo

                    });

                    Object Loadfila = Load[0];

                    SubirInfo(Loadfila, Convert.ToString(IdGenerado));
                }
            }

            return "Carga Exitosas";
        }

        public string SubirInfo(Object Load, string IdGenerado)
        {
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

        public PuntoModel GetPunto(DataRow row)
        {
            List<PuntoModel> Punto = new List<PuntoModel>();
            Punto.Add(new PuntoModel
            {
                    entrega = (EntregaModel)GetEntrega(row)
                ,   recoleccion = (RecoleccionModel)GetRecoleccion(row)

            });

            return Punto[0];
        }

        public RemolqueModel GetRemolque(DataRow row)
        {
            List<RemolqueModel> Punto = new List<RemolqueModel>();
            Punto.Add(new RemolqueModel
            {
                    remolque1 = (RemolqueDetalleModel)GetDetalleRemolque(row)
                ,   remolque2 = (RemolqueDetalleModel)GetDetalleRemolque(row)

            });

            return Punto[0];
        }

        public ConfigModel GetConfig(DataRow row)
        {
            List<ConfigModel> Config = new List<ConfigModel>();
            Config.Add(new ConfigModel
            {
                config = (ConfigConfigModel)GetConfigConfig(row)
            });

            return Config[0];
        }

        public Object GetEntrega(DataRow row)
        {
            EntregaModel PEntrega = new EntregaModel();

            if (row.Table.Columns.Contains("record_id") && row["record_id"] is not null)
            {
                PEntrega.record_id = (string)row["record_id"];
            }

            if (row.Table.Columns.Contains("administrative_area") && row["administrative_area"] is not null)
            {
                PEntrega.administrative_area = (string)row["administrative_area"];
            }

            if (row.Table.Columns.Contains("country") && row["country"] is not null)
            {
                PEntrega.country = (string)row["country"];
            }

            if (row.Table.Columns.Contains("fecha") && row["fecha"] is not null)
            {
                PEntrega.fecha = (string)row["fecha"];
            }

            if (row.Table.Columns.Contains("hora") && row["hora"] is not null)
            {
                PEntrega.hora = (string)row["hora"];
            }

            if (row.Table.Columns.Contains("locality") && row["locality"] is not null)
            {
                PEntrega.locality = (string)row["locality"];
            }

            if (row.Table.Columns.Contains("latitud") && row["latitud"] is not null)
            {
                PEntrega.latitud = (decimal)row["latitud"];
            }

            if (row.Table.Columns.Contains("longitud") && row["longitud"] is not null)
            {
                PEntrega.longitud = (decimal)row["longitud"];
            }

            if (row.Table.Columns.Contains("postal_code") && row["postal_code"] is not null)
            {
                PEntrega.postal_code = (string)row["postal_code"];
            }

            if (row.Table.Columns.Contains("sublocality") && row["sublocality"] is not null)
            {
                PEntrega.sublocality = (string)row["sublocality"];
            }

            return PEntrega;
        }

        public Object GetRecoleccion(DataRow row)
        {
            RecoleccionModel PRecoleccion = new RecoleccionModel();

            if (row.Table.Columns.Contains("record_id") && row["record_id"] is not null)
            {
                PRecoleccion.record_id = (string)row["record_id"];
            }

            if (row.Table.Columns.Contains("administrative_area") && row["administrative_area"] is not null)
            {
                PRecoleccion.administrative_area = (string)row["administrative_area"];
            }

            if (row.Table.Columns.Contains("country") && row["country"] is not null)
            {
                PRecoleccion.country = (string)row["country"];
            }

            if (row.Table.Columns.Contains("fecha") && row["fecha"] is not null)
            {
                PRecoleccion.fecha = (string)row["fecha"];
            }

            if (row.Table.Columns.Contains("hora") && row["hora"] is not null)
            {
                PRecoleccion.hora = (string)row["hora"];
            }

            if (row.Table.Columns.Contains("locality") && row["locality"] is not null)
            {
                PRecoleccion.locality = (string)row["locality"];
            }

            if (row.Table.Columns.Contains("latitud") && row["latitud"] is not null)
            {
                PRecoleccion.latitud = (decimal)row["latitud"];
            }

            if (row.Table.Columns.Contains("longitud") && row["longitud"] is not null)
            {
                PRecoleccion.longitud = (decimal)row["longitud"];
            }

            if (row.Table.Columns.Contains("postal_code") && row["postal_code"] is not null)
            {
                PRecoleccion.postal_code = (string)row["postal_code"];
            }

            if (row.Table.Columns.Contains("sublocality") && row["sublocality"] is not null)
            {
                PRecoleccion.sublocality = (string)row["sublocality"];
            }

            return PRecoleccion;
        }

        public Object GetConfigConfig(DataRow row)
        {
            ConfigConfigModel CConfig = new ConfigConfigModel();

            if (row.Table.Columns.Contains("record_id") && row["record_id"] is not null)
            {
                CConfig.record_id = (string)row["record_id"];
            }

            if (row.Table.Columns.Contains("estatusCarga") && row["estatusCarga"] is not null)
            {
                CConfig.estatusCarga = (string)row["estatusCarga"];
            }

            if (row.Table.Columns.Contains("fechaActualizacion") && row["fechaActualizacion"] is not null)
            {
                CConfig.fechaActualizacion = (string)row["fechaActualizacion"];
            }

            if (row.Table.Columns.Contains("fechaCreado") && row["fechaCreado"] is not null)
            {
                CConfig.fechaCreado = (string)row["fechaCreado"];
            }

            if (row.Table.Columns.Contains("notificacionOferta") && row["notificacionOferta"] is not null)
            {
                CConfig.notificacionOferta = (Boolean)row["notificacionOferta"];
            }

            return CConfig;
        }

        public Object GetDetalleRemolque(DataRow row)
        {
            RemolqueDetalleModel DRemolque = new RemolqueDetalleModel();

            if (row.Table.Columns.Contains("record_id") && row["record_id"] is not null)
            {
                DRemolque.record_id = (string)row["record_id"];
            }

            if (row.Table.Columns.Contains("alto") && row["alto"] is not null)
            {
                DRemolque.alto = (string)row["alto"];
            }

            if (row.Table.Columns.Contains("ancho") && row["ancho"] is not null)
            {
                DRemolque.ancho = (string)row["ancho"];
            }

            if (row.Table.Columns.Contains("contenedorTamano") && row["contenedorTamano"] is not null)
            {
                DRemolque.contenedorTamano = (string)row["contenedorTamano"];
            }

            if (row.Table.Columns.Contains("contenedorTipo") && row["contenedorTipo"] is not null)
            {
                DRemolque.contenedorTipo = (string)row["contenedorTipo"];
            }

            if (row.Table.Columns.Contains("embalaje") && row["embalaje"] is not null)
            {
                DRemolque.embalaje = (string)row["embalaje"];
            }

            if (row.Table.Columns.Contains("largo") && row["largo"] is not null)
            {
                DRemolque.largo = (string)row["largo"];
            }

            if (row.Table.Columns.Contains("peso") && row["peso"] is not null)
            {
                DRemolque.peso = (decimal)row["peso"];
            }

            if (row.Table.Columns.Contains("piezas") && row["piezas"] is not null)
            {
                DRemolque.piezas = (string)row["piezas"];
            }

            if (row.Table.Columns.Contains("volumen") && row["volumen"] is not null)
            {
                DRemolque.volumen = (decimal)row["peso"];
            }

            return DRemolque;
        }
    }
}
