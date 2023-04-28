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
using System.Timers;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml.Linq;
using NPOI.SS.Formula.Functions;
using NPOI.POIFS.FileSystem;
using static System.Net.WebRequestMethods;
using static Org.BouncyCastle.Math.EC.ECCurve;

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

        public List<DataLoadsModel> ProcesaExcel([FromForm] IFormFile ArchivoExcel, string NombreArchivo, ResponseModel response, string userConsig)
        {
            Stream stream = ArchivoExcel.OpenReadStream();
            List<DataLoadsModel> ListData = new List<DataLoadsModel>();
            List<DataLoadsModel> ListDataMaster = new List<DataLoadsModel>();

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

                if (cantidadFilas == 0)
                {
                    continue;
                }

                int intdtcolumn = HojaExcel.GetRow(0).LastCellNum;
                DataTable dtExcelData = GeneraTabla(HojaExcel.GetRow(0), HojaExcel.GetRow(0).LastCellNum);

                for (int ifila = 0; ifila <= cantidadFilas; ifila++)
                {
                    IRow fila = HojaExcel.GetRow(ifila);

                    if (ifila == 0)
                    {
                        continue;
                    }

                    InsertaTabla(dtExcelData, fila, intdtcolumn);
                }

                ListData = GeneraLoad(dtExcelData, NombreArchivo, userConsig);

                foreach (var item in ListData)
                {
                    ListDataMaster.Add(item);
                }
            }

            return ListDataMaster;
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
                if (dr is not null)
                {
                    if (dr.GetCell(icolumn) == null)
                    {
                        renglon[icolumn] = "";
                    }
                    else
                    {
                        renglon[icolumn] = dr.GetCell(icolumn).ToString();
                    }
                }
            }

            dtExcelData.Rows.Add(renglon);

            return dtExcelData;
        }

        public List<DataLoadsModel> GeneraLoad(DataTable dtExcelData, string NombreArchivo, string userConsig)
        {
            var message = "";
            List<DataLoadsModel> ListData = new List<DataLoadsModel>();
            DataLoadsModel dataLoads = new DataLoadsModel();

            if (dtExcelData.Rows.Count > 0)
            {
                foreach (DataRow row in dtExcelData.Rows)
                {
                    string cargaDescripcion;
                    int contador_error = 0;
                    List<LoadsModel> Loads = new List<LoadsModel>();

                    LoadsModel load = new LoadsModel();

                    Guid IdGenerado = Guid.NewGuid();

                    load.IdLoad = Convert.ToString(IdGenerado);

                    if (row.Table.Columns.Contains("cargaDescripcion") && row.Field<string>("cargaDescripcion") is not null && row.Field<string>("cargaDescripcion") != "")
                    {
                        load.cargaDescripcion = (string)row["cargaDescripcion"];
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("cargaRefrigerada") && row.Field<string>("cargaRefrigerada") is not null && row.Field<string>("cargaRefrigerada") != "")
                    {
                        load.cargaRefrigerada = Convert.ToBoolean((string)row["cargaRefrigerada"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("cargaTitulo") && row.Field<string>("cargaTitulo") is not null && row.Field<string>("cargaTitulo") != "")
                    {
                        load.cargaTitulo = (string)row["cargaTitulo"];
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("numRemolques") && row.Field<string>("numRemolques") is not null && row.Field<string>("numRemolques") != "")
                    {
                        load.numRemolques = (string)row["numRemolques"];
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("precioViaje") && row.Field<string>("precioViaje") is not null && row.Field<string>("precioViaje") != "")
                    {
                        load.precioViaje = Convert.ToDecimal((string)row["precioViaje"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("recibirOfertas") && row.Field<string>("recibirOfertas") is not null && row.Field<string>("recibirOfertas") != "")
                    {
                        load.recibirOfertas = Convert.ToBoolean((string)row["recibirOfertas"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("recomenEstibar") && row.Field<string>("recomenEstibar") is not null && row.Field<string>("recomenEstibar") != "")
                    {
                        load.recomenEstibar = Convert.ToBoolean((string)row["recomenEstibar"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("recomenFragil") && row.Field<string>("recomenFragil") is not null && row.Field<string>("recomenFragil") != "")
                    {
                        load.recomenFragil = Convert.ToBoolean((string)row["recomenFragil"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("recomenManejoCuidado") && row.Field<string>("recomenManejoCuidado") is not null && row.Field<string>("recomenManejoCuidado") != "")
                    {
                        load.recomenManejoCuidado = Convert.ToBoolean((string)row["recomenManejoCuidado"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("recomenMantenerSeco") && row.Field<string>("recomenMantenerSeco") is not null && row.Field<string>("recomenMantenerSeco") != "")
                    {
                        load.recomenMantenerSeco = Convert.ToBoolean((string)row["recomenMantenerSeco"]);
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("tipoCarga") && row.Field<string>("tipoCarga") is not null && row.Field<string>("tipoCarga") != "")
                    {
                        load.tipoCarga = (string)row["tipoCarga"];
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if (row.Table.Columns.Contains("tipoUnidad") && row.Field<string>("tipoUnidad") is not null && row.Field<string>("tipoUnidad") != "")
                    {
                        load.tipoUnidad = (string)row["tipoUnidad"];
                    }
                    else
                    {
                        contador_error += 1;
                    }

                    if(contador_error > 8)
                    {
                        continue;
                    }

                    load.userConsig = userConsig;
                    load.Punto = GetPunto(row);
                    load.Remolque = GetRemolque(row);
                    load.config = GetConfig(row);
                    load.nombreExcel = NombreArchivo;

                    Loads.Add(load);

                    Object Loadfila = Loads[0];

                    message = SubirInfo(Loadfila, Convert.ToString(IdGenerado));

                    if (message == "Id Cargado")
                    {
                        dataLoads.isSucces = true;
                    }
                    else
                    {
                        dataLoads.isSucces = false;
                    }

                    dataLoads.message = message;
                    dataLoads.idLoad = Convert.ToString(IdGenerado);
                    dataLoads.Load = Loads[0];

                    ListData.Add(dataLoads);
                }
            }

            return ListData;
        }

        public string SubirInfo(Object Load, string IdGenerado)
        {
            SetResponse response = cliente.Set("projects/proj_meqjHnqVDFjzhizHdj6Fjq/data/Loads/" + IdGenerado, Load);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return "Id Cargado";
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
                    entrega = (PuntoDetalleModel)GetEntrega(row)
                ,   recoleccion = (PuntoDetalleModel)GetRecoleccion(row)

            });

            return Punto[0];
        }

        public RemolqueModel GetRemolque(DataRow row)
        {
            List<RemolqueModel> Punto = new List<RemolqueModel>();
            Punto.Add(new RemolqueModel
            {
                    uno = (RemolqueDetalleModel)GetDetalleRemolque1(row)
                ,   dos = (RemolqueDetalleModel)GetDetalleRemolque2(row)

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
            PuntoDetalleModel PEntrega = new PuntoDetalleModel();

            if (row.Table.Columns.Contains("entrega_record_id") && row["entrega_record_id"] is not null)
            {
                PEntrega.record_id = (string)row["entrega_record_id"];
            }

            if (row.Table.Columns.Contains("entrega_address") && row["entrega_address"] is not null)
            {
                PEntrega.address = (string)row["entrega_address"];

                PEntrega.location = (LocationModel)ObtenerDatosRuta(PEntrega.address);
            }

            if (row.Table.Columns.Contains("entrega_administrative_area") && row["entrega_administrative_area"] is not null)
            {
                PEntrega.administrative_area = (string)row["entrega_administrative_area"];
            }

            if (row.Table.Columns.Contains("entrega_country") && row["entrega_country"] is not null)
            {
                PEntrega.country = (string)row["entrega_country"];
            }

            if (row.Table.Columns.Contains("entrega_fecha") && row["entrega_fecha"] is not null)
            {
                PEntrega.fecha = (string)row["entrega_fecha"];
            }

            if (row.Table.Columns.Contains("entrega_hora") && row["entrega_hora"] is not null)
            {
                PEntrega.hora = (string)row["entrega_hora"];
            }

            if (row.Table.Columns.Contains("entrega_locality") && row["entrega_locality"] is not null)
            {
                PEntrega.locality = (string)row["entrega_locality"];
            }

            if (row.Table.Columns.Contains("entrega_postal_code") && row["entrega_postal_code"] is not null)
            {
                PEntrega.postal_code = (string)row["entrega_postal_code"];
            }

            if (row.Table.Columns.Contains("entrega_sublocality") && row["entrega_sublocality"] is not null)
            {
                PEntrega.sublocality = (string)row["entrega_sublocality"];
            }

            return PEntrega;
        }

        public Object GetRecoleccion(DataRow row)
        {
            PuntoDetalleModel PRecoleccion = new PuntoDetalleModel();
            LocationModel location = new LocationModel();

            if (row.Table.Columns.Contains("recoleccion_record_id") && row["recoleccion_record_id"] is not null)
            {
                PRecoleccion.record_id = (string)row["recoleccion_record_id"];
            }

            if (row.Table.Columns.Contains("recoleccion_address") && row["recoleccion_address"] is not null)
            {
                PRecoleccion.address = (string)row["recoleccion_address"];

                PRecoleccion.location = (LocationModel)ObtenerDatosRuta(PRecoleccion.address);
            }

            if (row.Table.Columns.Contains("recoleccion_administrative_area") && row["recoleccion_administrative_area"] is not null)
            {
                PRecoleccion.administrative_area = (string)row["recoleccion_administrative_area"];
            }

            if (row.Table.Columns.Contains("recoleccion_country") && row["recoleccion_country"] is not null)
            {
                PRecoleccion.country = (string)row["recoleccion_country"];
            }

            if (row.Table.Columns.Contains("recoleccion_fecha") && row["recoleccion_fecha"] is not null)
            {
                PRecoleccion.fecha = (string)row["recoleccion_fecha"];
            }

            if (row.Table.Columns.Contains("recoleccion_hora") && row["recoleccion_hora"] is not null)
            {
                PRecoleccion.hora = (string)row["recoleccion_hora"];
            }

            if (row.Table.Columns.Contains("recoleccion_locality") && row["recoleccion_locality"] is not null)
            {
                PRecoleccion.locality = (string)row["recoleccion_locality"];
            }

            if (row.Table.Columns.Contains("recoleccion_postal_code") && row["recoleccion_postal_code"] is not null)
            {
                PRecoleccion.postal_code = (string)row["recoleccion_postal_code"];
            }

            if (row.Table.Columns.Contains("recoleccion_sublocality") && row["recoleccion_sublocality"] is not null)
            {
                PRecoleccion.sublocality = (string)row["recoleccion_sublocality"];
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

        public Object GetDetalleRemolque1(DataRow row)
        {
            RemolqueDetalleModel DRemolque = new RemolqueDetalleModel();

            if (row.Table.Columns.Contains("rem1_record_id") && row["rem1_record_id"] is not null)
            {
                DRemolque.record_id = (string)row["rem1_record_id"];
            }

            if (row.Table.Columns.Contains("rem1_alto") && row["rem1_alto"] is not null)
            {
                DRemolque.alto = (string)row["rem1_alto"];
            }

            if (row.Table.Columns.Contains("rem1_ancho") && row["rem1_ancho"] is not null)
            {
                DRemolque.ancho = (string)row["rem1_ancho"];
            }

            if (row.Table.Columns.Contains("rem1_contenedorTamano") && row["rem1_contenedorTamano"] is not null)
            {
                DRemolque.contenedorTamano = (string)row["rem1_contenedorTamano"];
            }

            if (row.Table.Columns.Contains("rem1_contenedorTipo") && row["rem1_contenedorTipo"] is not null)
            {
                DRemolque.contenedorTipo = (string)row["rem1_contenedorTipo"];
            }

            if (row.Table.Columns.Contains("rem1_embalaje") && row["rem1_embalaje"] is not null)
            {
                DRemolque.embalaje = (string)row["rem1_embalaje"];
            }

            if (row.Table.Columns.Contains("rem1_largo") && row["rem1_largo"] is not null)
            {
                DRemolque.largo = (string)row["rem1_largo"];
            }

            if (row.Table.Columns.Contains("rem1_peso") && row["rem1_peso"] is not null)
            {
                DRemolque.peso = (string)row["rem1_peso"];
            }

            if (row.Table.Columns.Contains("rem1_piezas") && row["rem1_piezas"] is not null)
            {
                DRemolque.piezas = (string)row["rem1_piezas"];
            }

            if (row.Table.Columns.Contains("rem1_volumen") && row["rem1_volumen"] is not null)
            {
                DRemolque.volumen = (string)row["rem1_volumen"];
            }

            return DRemolque;
        }

        public Object GetDetalleRemolque2(DataRow row)
        {
            RemolqueDetalleModel DRemolque = new RemolqueDetalleModel();

            if (row.Table.Columns.Contains("rem2_record_id") && row["rem2_record_id"] is not null)
            {
                DRemolque.record_id = (string)row["rem2_record_id"];
            }

            if (row.Table.Columns.Contains("rem2_alto") && row["rem2_alto"] is not null)
            {
                DRemolque.alto = (string)row["rem2_alto"];
            }

            if (row.Table.Columns.Contains("rem2_ancho") && row["rem2_ancho"] is not null)
            {
                DRemolque.ancho = (string)row["rem2_ancho"];
            }

            if (row.Table.Columns.Contains("rem2_contenedorTamano") && row["rem2_contenedorTamano"] is not null)
            {
                DRemolque.contenedorTamano = (string)row["rem2_contenedorTamano"];
            }

            if (row.Table.Columns.Contains("rem2_contenedorTipo") && row["rem2_contenedorTipo"] is not null)
            {
                DRemolque.contenedorTipo = (string)row["rem2_contenedorTipo"];
            }

            if (row.Table.Columns.Contains("rem2_embalaje") && row["rem2_embalaje"] is not null)
            {
                DRemolque.embalaje = (string)row["rem2_embalaje"];
            }

            if (row.Table.Columns.Contains("rem2_largo") && row["rem2_largo"] is not null)
            {
                DRemolque.largo = (string)row["rem2_largo"];
            }

            if (row.Table.Columns.Contains("rem2_peso") && row["rem2_peso"] is not null)
            {
                DRemolque.peso = (string)row["rem2_peso"];
            }

            if (row.Table.Columns.Contains("rem2_piezas") && row["rem2_piezas"] is not null)
            {
                DRemolque.piezas = (string)row["rem2_piezas"];
            }

            if (row.Table.Columns.Contains("rem2_volumen") && row["rem2_volumen"] is not null)
            {
                DRemolque.volumen = (string)row["rem2_volumen"];
            }

            return DRemolque;
        }

        //FUNCION PARA OBTENER DATOS DE LOS PUNTOS DE ENTREGA Y RECOLECCION
        public Object ObtenerDatosRuta(string address)
        {
            LocationModel location = new LocationModel(); 

            var apikey = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

            string requestUri = string.Format("https://maps.googleapis.com/maps/api/geocode/xml?key={1}&address={0}&sensor=false", Uri.EscapeDataString(address), apikey);

            WebRequest request = WebRequest.Create(requestUri);
            WebResponse response = request.GetResponse();
            XDocument xdoc = XDocument.Load(response.GetResponseStream());

            XElement result = xdoc.Element("GeocodeResponse").Element("result");
            XElement locationElement = result.Element("geometry").Element("location");
            XElement lat = locationElement.Element("lat");
            XElement lng = locationElement.Element("lng");

            location.latitud = (decimal)lat;
            location.longitud = (decimal)lng;

            return location;
        }

    }
}
