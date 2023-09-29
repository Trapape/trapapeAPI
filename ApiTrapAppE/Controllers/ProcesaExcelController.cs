using ApiTrapAppE.Models;
using FireSharp;
using FireSharp.Config;
using FireSharp.Interfaces;
using FireSharp.Response;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ApiTrapAppE.Controllers
{
    public class ProcesaExcelController : Controller
    {
        private IFirebaseClient cliente;

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

        public List<DataLoadsModel> ProcesaExcel([FromForm] IFormFile ArchivoExcel, string NombreArchivo, string userConsig, string downloadURL, string idCargaPrincipal)
        {
            // Inicialización de variables
            Stream stream = ArchivoExcel.OpenReadStream();
            List<DataLoadsModel> ListDataMaster = new List<DataLoadsModel>();
            IWorkbook MiExcel = null;

            // Determinar el tipo de archivo Excel (.xlsx o .xls) y crear el objeto adecuado
            if (Path.GetExtension(NombreArchivo) == ".xlsx")
                MiExcel = new XSSFWorkbook(stream);
            else
                MiExcel = new HSSFWorkbook(stream);

            int cantidadHojas = MiExcel.NumberOfSheets; // Obtener el número de hojas en el archivo Excel

            // Recorrer las hojas para procesar cada carga
            for (int hoja = 0; hoja < cantidadHojas; hoja++)
            {
                ISheet HojaExcel = MiExcel.GetSheetAt(hoja);

                if (!HojaExcel.SheetName.Contains("_"))
                    continue; // Saltar hojas que no contienen cierto carácter especial

                DataTable dtExcelData = null;

                // Recorrer las filas de la hoja
                for (int ifila = 0; ifila <= HojaExcel.LastRowNum; ifila++)
                {
                    IRow fila = HojaExcel.GetRow(ifila);

                    if (fila == null)
                        continue; // Saltar filas nulas

                    if (ifila == 0)
                    {
                        // La primera fila (cabecera) se usa para crear la tabla solo una vez
                        int intdtcolumn = fila.LastCellNum;
                        dtExcelData = GeneraTabla(fila, intdtcolumn);
                        continue;
                    }

                    if (fila.GetCell(0) == null || fila.GetCell(0).CellType == CellType.Blank || (fila.GetCell(0).StringCellValue != "true" && fila.GetCell(0).StringCellValue != "false"))
                    {
                        // Si la primera celda está vacía, se detiene el procesamiento de filas en esta hoja
                        break;
                    }

                    int cabecera = dtExcelData.Columns.Count - 5;
                    if (fila.LastCellNum > cabecera)
                        InsertaTabla(dtExcelData, fila, cabecera);
                }

                if (dtExcelData != null && dtExcelData.Rows.Count > 0)
                {
                    // Eliminar filas vacías de dtExcelData
                    dtExcelData = dtExcelData.Rows.Cast<DataRow>()
                                    .Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field.ToString())))
                                    .CopyToDataTable();

                    // Generar datos de carga y agregarlos a la lista maestra solo si se creó la tabla
                    List<DataLoadsModel> ListData = GeneraLoad(dtExcelData, NombreArchivo, userConsig, downloadURL, idCargaPrincipal);
                    ListDataMaster.AddRange(ListData);
                }
            }

            return ListDataMaster;
        }

        public DataTable GeneraTabla(IRow dr, int intdtcolumn)
        {
            var dtExcelData = new DataTable();

            for (int icolumn = 0; icolumn < intdtcolumn; icolumn++)
            {
                ICell cell = dr.GetCell(icolumn);

                if (cell != null)
                {
                    string columnName = cell.ToString().Trim(); // Eliminar espacios en blanco alrededor del nombre de la columna

                    // Verificar si el nombre de la columna ya existe en la tabla
                    if (!dtExcelData.Columns.Contains(columnName))
                    {
                        dtExcelData.Columns.Add(columnName);
                    }
                    else
                    {
                        // Si el nombre de la columna ya existe, puedes modificarlo para hacerlo único
                        int count = 1;
                        while (dtExcelData.Columns.Contains(columnName + "_" + count))
                        {
                            count++;
                        }
                        columnName = columnName + "_" + count;
                        dtExcelData.Columns.Add(columnName);
                    }
                }
            }

            return dtExcelData;
        }

        public DataTable InsertaTabla(DataTable dtExcelData, IRow dr, int intdtcolumn)
        {
            // Crear un nuevo DataRow para agregar a la tabla
            DataRow renglon = dtExcelData.NewRow();

            for (int icolumn = 0; icolumn < intdtcolumn; icolumn++)
            {
                if (dr is not null)
                {
                    if (dr.GetCell(icolumn) == null)
                    {
                        // Si la celda es nula, asigna una cadena vacía
                        renglon[icolumn] = "";
                    }
                    else
                    {
                        // Obtener el valor de la primera celda
                        string valor_primer_celda = dr.GetCell(0).StringCellValue;

                        // Verificar si la primera celda está vacía
                        if (valor_primer_celda == "")
                        {
                            // Si la primera celda está vacía, se detiene el procesamiento de esta fila
                            break;
                        }

                        // Determinar el tipo de dato de la celda actual
                        string tipo_dato_celda = dr.GetCell(icolumn).CachedFormulaResultType.ToString();

                        // Asignar el valor de la celda actual al DataRow según su tipo de dato
                        if (tipo_dato_celda == "String")
                        {
                            // Si el tipo de dato es "String", asigna el valor como cadena
                            renglon[icolumn] = dr.GetCell(icolumn).StringCellValue;
                        }
                        else if (tipo_dato_celda == "Numeric")
                        {
                            // Si el tipo de dato es "Numeric", asigna el valor numérico convertido a cadena
                            renglon[icolumn] = dr.GetCell(icolumn).NumericCellValue.ToString();
                        }
                        // Puedes agregar más casos aquí para otros tipos de datos si es necesario
                    }
                }
            }

            // Agregar el DataRow a la tabla
            dtExcelData.Rows.Add(renglon);

            // Devolver la tabla actualizada con la nueva fila
            return dtExcelData;
        }

        public List<DataLoadsModel> GeneraLoad(DataTable dtExcelData, string NombreArchivo, string userConsig, string downloadURL, string idCargaPrincipal)
        {
            var ListData = new List<DataLoadsModel>();

            // Verifica si la tabla contiene datos
            if (dtExcelData.Rows.Count > 0)
            {
                foreach (DataRow row in dtExcelData.Rows)
                {
                    int contador_error = 0;
                    List<LoadsModel> Loads = new List<LoadsModel>();
                    LoadsModel load = new LoadsModel();
                    Guid IdGenerado = Guid.NewGuid();

                    // Asigna un GUID como el IdLoad
                    load.IdLoad = IdGenerado.ToString();

                    if (row.Table.Columns.Contains("cargaDescripcion") && !string.IsNullOrEmpty(row.Field<string>("cargaDescripcion")))
                        load.cargaDescripcion = (string)row["cargaDescripcion"];
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("cargaRefrigerada") && !string.IsNullOrEmpty(row.Field<string>("cargaRefrigerada")))
                        load.cargaRefrigerada = Convert.ToBoolean((string)row["cargaRefrigerada"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("cargaTitulo") && !string.IsNullOrEmpty(row.Field<string>("cargaTitulo")))
                        load.cargaTitulo = (string)row["cargaTitulo"];
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("numRemolques") && !string.IsNullOrEmpty(row.Field<string>("numRemolques")))
                        load.numRemolques = (string)row["numRemolques"];
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("precioViaje") && !string.IsNullOrEmpty(row.Field<string>("precioViaje")))
                        load.precioViaje = Convert.ToDecimal((string)row["precioViaje"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("recibirOfertas") && !string.IsNullOrEmpty(row.Field<string>("recibirOfertas")))
                        load.recibirOfertas = Convert.ToBoolean((string)row["recibirOfertas"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("recomenEstibar") && !string.IsNullOrEmpty(row.Field<string>("recomenEstibar")))
                        load.recomenEstibar = Convert.ToBoolean((string)row["recomenEstibar"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("recomenFragil") && !string.IsNullOrEmpty(row.Field<string>("recomenFragil")))
                        load.recomenFragil = Convert.ToBoolean((string)row["recomenFragil"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("recomenManejoCuidado") && !string.IsNullOrEmpty(row.Field<string>("recomenManejoCuidado")))
                        load.recomenManejoCuidado = Convert.ToBoolean((string)row["recomenManejoCuidado"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("recomenMantenerSeco") && !string.IsNullOrEmpty(row.Field<string>("recomenMantenerSeco")))
                        load.recomenMantenerSeco = Convert.ToBoolean((string)row["recomenMantenerSeco"]);
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("tipoCarga") && !string.IsNullOrEmpty(row.Field<string>("tipoCarga")))
                        load.tipoCarga = (string)row["tipoCarga"];
                    else
                        contador_error += 1;

                    if (row.Table.Columns.Contains("tipoUnidad") && !string.IsNullOrEmpty(row.Field<string>("tipoUnidad")))
                        load.tipoUnidad = (string)row["tipoUnidad"];
                    else
                        contador_error += 1;

                    // Configura otras propiedades de LoadsModel
                    load.userConsig = userConsig;
                    load.Punto = GetPunto(row);
                    load.remolque = GetRemolque(row);
                    load.config = GetConfig(row);
                    load.nombreExcel = downloadURL;
                    load.idCargaPrincipal = idCargaPrincipal;

                    // Calcula la distancia y el tiempo si no están disponibles
                    decimal latitude_origen = load.Punto.recoleccion.location.latitude;
                    decimal longitude_origen = load.Punto.recoleccion.location.longitude;
                    decimal latitude_destino = load.Punto.entrega.location.latitude;
                    decimal longitude_destino = load.Punto.entrega.location.longitude;
                    string distancia_tiempo = Calcula_Time_KM(latitude_origen, longitude_origen, latitude_destino, longitude_destino);
                    int distanciaMetros = Convert.ToInt32(distancia_tiempo.Split(",")[0]);
                    load.distanciaKM = Convert.ToDecimal(distanciaMetros / 1000);
                    int duration = Convert.ToInt32(distancia_tiempo.Split(",")[1]);
                    load.tiempoRuta = Convert.ToString(Math.Round((duration * 2.3), 2));

                    // Agrega el objeto Load a la lista Loads
                    Loads.Add(load);

                    // Obtiene el primer objeto Load de la lista
                    Object Loadfila = Loads[0];

                    // Crea una nueva instancia de DataLoadsModel
                    DataLoadsModel dataLoads = new DataLoadsModel();

                    if (row.Table.Columns.Contains("cargaValida") && row.Field<string>("cargaValida") is not null && row.Field<string>("cargaValida") == "true")
                    {
                        dataLoads.isSucces = true;
                        dataLoads.message = SubirInfo(Loadfila, IdGenerado.ToString());

                        // Si hay demasiados errores, pasa a la siguiente fila
                        if (contador_error > 8 || dataLoads.message != "Id Cargado")
                            dataLoads.isSucces = false;
                    }
                    else
                    {
                        //message = "No se cargaron los datos.";
                        dataLoads.isSucces = false;
                        dataLoads.message = "No se cargaron los datos.";
                    }

                    dataLoads.idLoad = IdGenerado.ToString();
                    dataLoads.Load = Loads[0];

                    // Agrega el objeto DataLoadsModel a la lista final
                    ListData.Add(dataLoads);
                }
            }

            // Devuelve la lista de DataLoadsModel
            return ListData;
        }

        public string SubirInfo(Object Load, string IdGenerado)
        {
            SetResponse response = cliente.Set("projects/proj_meqjHnqVDFjzhizHdj6Fjq/data/Loads/" + IdGenerado, Load);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                _ = SubirGeoFireAsync(Load, IdGenerado);
                return "Id Cargado";
            }
            else
            {
                return "No se cargaron los datos.";
            }
        }

        public async Task<string> SubirGeoFireAsync(object load, string idGenerado)
        {
            // Obtener los valores de latitud y longitud
            LoadsModel aux = (LoadsModel)load;

            // Construir la URL con los parámetros individuales
            string url = $"https://us-central1-trapape.cloudfunctions.net/appProcessWebHook/ccp_hBWXSJ5GAN9Djb6NK1LaHB?origen_lat={aux.Punto.recoleccion.location.latitude}&origen_lon={aux.Punto.recoleccion.location.longitude}&destino_lat={aux.Punto.entrega.location.latitude}&destino_lon={aux.Punto.entrega.location.longitude}&idLoad={idGenerado}";

            // Crear una instancia de HttpClient
            using var httpClient = new HttpClient();

            try
            {
                // Realizar una solicitud HTTP GET a la URL construida
                var response = await httpClient.GetAsync(url);

                // Verificar si la solicitud fue exitosa (código de respuesta 200)
                if (response.IsSuccessStatusCode)
                {
                    // Leer la respuesta
                    string contenidoRespuesta = await response.Content.ReadAsStringAsync();
                    Console.WriteLine("Respuesta del servicio: " + contenidoRespuesta);
                    await Task.Delay(3000);
                    SetResponse responseData = cliente.Set("projects/proj_meqjHnqVDFjzhizHdj6Fjq/geoFireGroups/Loads/" + idGenerado + "/data", load);
                    if (responseData.StatusCode == System.Net.HttpStatusCode.OK) return "Id Cargado";
                    else return "No se cargaron los datos.";
                }
                else
                {
                    Console.WriteLine("La solicitud falló con código de respuesta: " + response.StatusCode);
                    return "No se cargaron los datos.";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al realizar la solicitud: " + ex.Message);
                return "No se cargaron los datos.";
            }
        }

        public PuntoModel GetPunto(DataRow row)
        {
            List<PuntoModel> Punto = new List<PuntoModel>();
            Punto.Add(new PuntoModel
            {
                entrega = (PuntoDetalleModel)GetEntrega(row)
                ,
                recoleccion = (PuntoDetalleModel)GetRecoleccion(row)
            });

            return Punto[0];
        }

        public RemolqueModel GetRemolque(DataRow row)
        {
            List<RemolqueModel> Punto = new List<RemolqueModel>();
            Punto.Add(new RemolqueModel
            {
                uno = (RemolqueDetalleModel)GetDetalleRemolque1(row)
                ,
                dos = (RemolqueDetalleModel)GetDetalleRemolque2(row)
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

        public PuntoDetalleModel GetEntrega(DataRow row)
        {
            PuntoDetalleModel PEntrega = new PuntoDetalleModel();
            LocationModel location = new LocationModel();
            XDocument xmladdress = new XDocument();
            string administrative_area = "", country = "", locality = "", postal_code = "", sublocality = "";

            if (row.Table.Columns.Contains("entrega_record_id") && row["entrega_record_id"] is not null)
                PEntrega.record_id = (string)row["entrega_record_id"];

            if (row.Table.Columns.Contains("entrega_address") && row["entrega_address"] is not null)
            {
                PEntrega.address = (string)row["entrega_address"];

                xmladdress = ObtenerDatosRuta(PEntrega.address);

                XElement result = xmladdress.Element("GeocodeResponse").Element("result");

                var xmlDocument = new XmlDocument();
                using (var xmlReader = result.CreateReader())
                {
                    xmlDocument.Load(xmlReader);
                }

                //DATOS ADDRESS
                XmlNodeList address_component = xmlDocument.SelectNodes("//address_component");

                foreach (XmlNode comp in address_component)
                {
                    string jsonText = JsonConvert.SerializeXmlNode(comp);

                    jsonText = jsonText.Replace("{\"address_component\":", "");
                    jsonText = jsonText.Replace("}}", "}");

                    if (jsonText.Contains("administrative_area"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        administrative_area = (string)property.Value;
                    }

                    if (jsonText.Contains("country"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        country = (string)property.Value;
                    }

                    if (jsonText.Contains("locality"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        locality = (string)property.Value;
                    }

                    if (jsonText.Contains("postal_code"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        postal_code = (string)property.Value;
                    }

                    if (jsonText.Contains("sublocality"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        sublocality = (string)property.Value;
                    }
                }

                //DATOS LOCATION
                XElement locationElement = result.Element("geometry").Element("location");
                XElement lat = locationElement.Element("lat");
                XElement lng = locationElement.Element("lng");

                location.latitude = (decimal)lat;
                location.longitude = (decimal)lng;

                PEntrega.location = location;
            }

            if (row.Table.Columns.Contains("entrega_administrative_area") && row["entrega_administrative_area"] is not null)
            {
                PEntrega.administrative_area = (string)row["entrega_administrative_area"];
                PEntrega.administrative_area_level_1 = (string)row["entrega_administrative_area"];
            }
            else
            {
                PEntrega.administrative_area = administrative_area;
                PEntrega.administrative_area_level_1 = administrative_area;
            }

            if (row.Table.Columns.Contains("entrega_country") && row["entrega_country"] is not null)
                PEntrega.country = (string)row["entrega_country"];
            else
                PEntrega.country = country;

            if (row.Table.Columns.Contains("entrega_fecha") && row["entrega_fecha"] is not null)
            {
                double fechaExcelValue = Convert.ToDouble(row["entrega_fecha"]);
                DateTime fechaEntrega = DateTime.FromOADate(fechaExcelValue);
                string fechaEntregaString = fechaEntrega.ToString("dd-MM-yyyy");
                PEntrega.fecha = fechaEntregaString;
                //PEntrega.fecha = (string)row["entrega_fecha"];
            }

            if (row.Table.Columns.Contains("entrega_hora") && row["entrega_hora"] is not null)
            {
                double horaExcelValue = Convert.ToDouble(row["entrega_hora"]);
                double horas = horaExcelValue * 24;
                TimeSpan horaEntrega = TimeSpan.FromHours(horas);
                string horaEntregaString = horaEntrega.ToString();
                PEntrega.hora = horaEntregaString;
                //PEntrega.hora = (string)row["entrega_hora"];
            }

            if (row.Table.Columns.Contains("entrega_locality") && row["entrega_locality"] is not null)

                PEntrega.locality = (string)row["entrega_locality"];
            else
                PEntrega.locality = locality;

            if (row.Table.Columns.Contains("entrega_postal_code") && row["entrega_postal_code"] is not null)

                PEntrega.postal_code = (string)row["entrega_postal_code"];
            else
                PEntrega.postal_code = postal_code;

            if (row.Table.Columns.Contains("entrega_sublocality") && row["entrega_sublocality"] is not null)

                PEntrega.sublocality = (string)row["entrega_sublocality"];
            else
                PEntrega.sublocality = sublocality;

            return PEntrega;
        }

        public Object GetRecoleccion(DataRow row)
        {
            PuntoDetalleModel PRecoleccion = new PuntoDetalleModel();
            LocationModel location = new LocationModel();
            XDocument xmladdress = new XDocument();
            string administrative_area = "", country = "", locality = "", postal_code = "", sublocality = "";

            if (row.Table.Columns.Contains("recoleccion_record_id") && row["recoleccion_record_id"] is not null)
            {
                PRecoleccion.record_id = (string)row["recoleccion_record_id"];
            }

            if (row.Table.Columns.Contains("recoleccion_address") && row["recoleccion_address"] is not null)
            {
                PRecoleccion.address = (string)row["recoleccion_address"];

                xmladdress = ObtenerDatosRuta(PRecoleccion.address);

                XElement result = xmladdress.Element("GeocodeResponse").Element("result");

                var xmlDocument = new XmlDocument();
                using (var xmlReader = result.CreateReader())
                {
                    xmlDocument.Load(xmlReader);
                }

                //DATOS ADDRESS
                XmlNodeList address_component = xmlDocument.SelectNodes("//address_component");

                foreach (XmlNode comp in address_component)
                {
                    string jsonText = JsonConvert.SerializeXmlNode(comp);

                    jsonText = jsonText.Replace("{\"address_component\":", "");
                    jsonText = jsonText.Replace("}}", "}");

                    if (jsonText.Contains("administrative_area"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        administrative_area = (string)property.Value;
                    }

                    if (jsonText.Contains("country"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        country = (string)property.Value;
                    }

                    if (jsonText.Contains("locality"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        locality = (string)property.Value;
                    }

                    if (jsonText.Contains("postal_code"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        postal_code = (string)property.Value;
                    }

                    if (jsonText.Contains("sublocality"))
                    {
                        JObject root = JObject.Parse(jsonText);
                        JProperty property = (JProperty)root.First.Next;
                        sublocality = (string)property.Value;
                    }
                }

                //DATOS LOCATION
                XElement locationElement = result.Element("geometry").Element("location");
                XElement lat = locationElement.Element("lat");
                XElement lng = locationElement.Element("lng");

                location.latitude = (decimal)lat;
                location.longitude = (decimal)lng;

                PRecoleccion.location = location;
            }

            if (row.Table.Columns.Contains("recoleccion_administrative_area") && row["recoleccion_administrative_area"] is not null)
            {
                PRecoleccion.administrative_area = (string)row["recoleccion_administrative_area"];
                PRecoleccion.administrative_area_level_1 = (string)row["recoleccion_administrative_area"];
            }
            else
            {
                PRecoleccion.administrative_area = administrative_area;
                PRecoleccion.administrative_area_level_1 = administrative_area;
            }

            if (row.Table.Columns.Contains("recoleccion_country") && row["recoleccion_country"] is not null)

                PRecoleccion.country = (string)row["recoleccion_country"];
            else
                PRecoleccion.country = country;

            if (row.Table.Columns.Contains("recoleccion_fecha") && row["recoleccion_fecha"] is not null)
            {
                double fechaExcelValue = Convert.ToDouble(row["recoleccion_fecha"]);
                DateTime fechaEntrega = DateTime.FromOADate(fechaExcelValue);
                string fechaEntregaString = fechaEntrega.ToString("dd-MM-yyyy");
                PRecoleccion.fecha = fechaEntregaString;
                //PRecoleccion.fecha = (string)row["recoleccion_fecha"];
            }

            if (row.Table.Columns.Contains("recoleccion_hora") && row["recoleccion_hora"] is not null)
            {
                double horaExcelValue = Convert.ToDouble(row["recoleccion_hora"]);
                double horas = horaExcelValue * 24;
                TimeSpan horaEntrega = TimeSpan.FromHours(horas);
                string horaEntregaString = horaEntrega.ToString();
                PRecoleccion.hora = horaEntregaString;
                //PRecoleccion.hora = (string)row["recoleccion_hora"];
            }

            if (row.Table.Columns.Contains("recoleccion_locality") && row["recoleccion_locality"] is not null)

                PRecoleccion.locality = (string)row["recoleccion_locality"];
            else
                PRecoleccion.locality = locality;

            if (row.Table.Columns.Contains("recoleccion_postal_code") && row["recoleccion_postal_code"] is not null)

                PRecoleccion.postal_code = (string)row["recoleccion_postal_code"];
            else
                PRecoleccion.postal_code = postal_code;

            if (row.Table.Columns.Contains("recoleccion_sublocality") && row["recoleccion_sublocality"] is not null)

                PRecoleccion.sublocality = (string)row["recoleccion_sublocality"];
            else
                PRecoleccion.sublocality = sublocality;

            return PRecoleccion;
        }

        public Object GetConfigConfig(DataRow row)
        {
            ConfigConfigModel CConfig = new ConfigConfigModel();

            if (row.Table.Columns.Contains("record_id") && row["record_id"] is not null)
                CConfig.record_id = (string)row["record_id"];
            else
                CConfig.record_id = "config";

            CConfig.fechaCreado = GetTimestamp(DateTime.Now);

            if (row.Table.Columns.Contains("notificacionOferta") && row["notificacionOferta"] is not null)
                CConfig.notificacionOferta = (Boolean)row["notificacionOferta"];
            else
                CConfig.notificacionOferta = false;

            return CConfig;
        }

        public Object GetDetalleRemolque1(DataRow row)
        {
            RemolqueDetalleModel DRemolque = new RemolqueDetalleModel();

            if (row.Table.Columns.Contains("rem1_record_id") && row["rem1_record_id"] is not null)
                DRemolque.record_id = (string)row["rem1_record_id"];

            if (row.Table.Columns.Contains("rem1_alto") && row["rem1_alto"] is not null)
                DRemolque.alto = (string)row["rem1_alto"];

            if (row.Table.Columns.Contains("rem1_ancho") && row["rem1_ancho"] is not null)
                DRemolque.ancho = (string)row["rem1_ancho"];

            if (row.Table.Columns.Contains("rem1_contenedorTamano") && row["rem1_contenedorTamano"] is not null)
                DRemolque.contenedorTamano = (string)row["rem1_contenedorTamano"];

            if (row.Table.Columns.Contains("rem1_contenedorTipo") && row["rem1_contenedorTipo"] is not null)
                DRemolque.contenedorTipo = (string)row["rem1_contenedorTipo"];

            if (row.Table.Columns.Contains("rem1_embalaje") && row["rem1_embalaje"] is not null)
                DRemolque.embalaje = (string)row["rem1_embalaje"];

            if (row.Table.Columns.Contains("rem1_largo") && row["rem1_largo"] is not null)
                DRemolque.largo = (string)row["rem1_largo"];

            if (row.Table.Columns.Contains("rem1_peso") && row["rem1_peso"] is not null)
                DRemolque.peso = (string)row["rem1_peso"];

            if (row.Table.Columns.Contains("rem1_piezas") && row["rem1_piezas"] is not null)
                DRemolque.piezas = (string)row["rem1_piezas"];

            if (row.Table.Columns.Contains("rem1_volumen") && row["rem1_volumen"] is not null)
                DRemolque.volumen = (string)row["rem1_volumen"];

            return DRemolque;
        }

        public Object GetDetalleRemolque2(DataRow row)
        {
            RemolqueDetalleModel DRemolque = new RemolqueDetalleModel();

            if (row.Table.Columns.Contains("rem2_record_id") && row["rem2_record_id"] is not null)
                DRemolque.record_id = (string)row["rem2_record_id"];

            if (row.Table.Columns.Contains("rem2_alto") && row["rem2_alto"] is not null)
                DRemolque.alto = (string)row["rem2_alto"];

            if (row.Table.Columns.Contains("rem2_ancho") && row["rem2_ancho"] is not null)
                DRemolque.ancho = (string)row["rem2_ancho"];

            if (row.Table.Columns.Contains("rem2_contenedorTamano") && row["rem2_contenedorTamano"] is not null)
                DRemolque.contenedorTamano = (string)row["rem2_contenedorTamano"];

            if (row.Table.Columns.Contains("rem2_contenedorTipo") && row["rem2_contenedorTipo"] is not null)
                DRemolque.contenedorTipo = (string)row["rem2_contenedorTipo"];

            if (row.Table.Columns.Contains("rem2_embalaje") && row["rem2_embalaje"] is not null)
                DRemolque.embalaje = (string)row["rem2_embalaje"];

            if (row.Table.Columns.Contains("rem2_largo") && row["rem2_largo"] is not null)
                DRemolque.largo = (string)row["rem2_largo"];

            if (row.Table.Columns.Contains("rem2_peso") && row["rem2_peso"] is not null)
                DRemolque.peso = (string)row["rem2_peso"];

            if (row.Table.Columns.Contains("rem2_piezas") && row["rem2_piezas"] is not null)
                DRemolque.piezas = (string)row["rem2_piezas"];

            if (row.Table.Columns.Contains("rem2_volumen") && row["rem2_volumen"] is not null)
                DRemolque.volumen = (string)row["rem2_volumen"];

            return DRemolque;
        }

        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy-MM-ddTHH:mm:ss");
        }

        //FUNCION PARA OBTENER DATOS DE LOS PUNTOS DE ENTREGA Y RECOLECCION
        public XDocument ObtenerDatosRuta(string address)
        {
            var apikey = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

            string requestUri = string.Format("https://maps.googleapis.com/maps/api/geocode/xml?key={1}&address={0}&sensor=false", Uri.EscapeDataString(address), apikey);

            WebRequest request = WebRequest.Create(requestUri);
            WebResponse response = request.GetResponse();

            var respmaps = response.GetResponseStream();

            XDocument xdoc = XDocument.Load(response.GetResponseStream());

            return xdoc;
        }

        public string Calcula_Time_KM(decimal latitude_origen, decimal longitude_origen, decimal latitude_destino, decimal longitude_destino)
        {
            var apikey = "AIzaSyBs-iRGy4GQdnqmLrDqMSV8sIcraM9kXl4";

            string requestUri = string.Format("https://maps.googleapis.com/maps/api/directions/json?origin=" + Convert.ToString(latitude_origen) + "," + Convert.ToString(longitude_origen) + "&destination=" + Convert.ToString(latitude_destino) + "," + Convert.ToString(longitude_destino) + "&mode=driving&lenguaje=es&key=" + apikey);

            WebRequest request = WebRequest.Create(requestUri);
            WebResponse response = request.GetResponse();

            var respmaps = response.GetResponseStream();

            Stream dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream, Encoding.UTF8);
            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            response.Close();

            string result = ProcesaResultApiGoogle(responseFromServer);

            return result;
        }

        public string ProcesaResultApiGoogle(string responseFromServer)
        {
            decimal distanciaFinal = 0;
            string strDataDistanciaJSON1 = "", strDataDistanciaJSON2 = "", resudistancia = "", strDataTiempoJSON = "", resutiempo = "";

            JObject jsonRespose = JObject.Parse(responseFromServer);
            List<JToken> jtDataResponse = new List<JToken>(jsonRespose.Children());
            object obj2 = jtDataResponse[1].First;

            string qwert = obj2.ToString();
            string nuevo = qwert.Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
            JArray nuevoarray = (JArray)JsonConvert.DeserializeObject(nuevo);
            JObject nuevoobj = (JObject)nuevoarray[0];

            int a = 0;
            IList<JToken> list = nuevoobj;
            for (int i = 0; i < list.Count; i++)
            {
                JToken item = list[i];
                if (a == 2)
                {
                    strDataDistanciaJSON1 = item.First.ToString();
                    strDataDistanciaJSON1 = strDataDistanciaJSON1.Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
                    break;
                }

                a = a + 1;
            }

            JArray nuevoarray2 = (JArray)JsonConvert.DeserializeObject(strDataDistanciaJSON1);
            JObject nuevoobj2 = (JObject)nuevoarray2[0];

            int b = 0;
            IList<JToken> list2 = nuevoobj2;
            for (int i = 0; i < list2.Count; i++)
            {
                JToken item = list2[i];
                if (b == 0)
                {
                    strDataDistanciaJSON2 = item.First.ToString();
                    strDataDistanciaJSON2 = strDataDistanciaJSON2.Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
                }

                if (b == 1)
                {
                    strDataTiempoJSON = item.First.ToString();
                    strDataTiempoJSON = strDataTiempoJSON.Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
                    break;
                }

                b = b + 1;
            }

            JObject resultadofinalObjectDistancia = (JObject)JsonConvert.DeserializeObject(strDataDistanciaJSON2);
            int c = 0;

            IList<JToken> list3 = resultadofinalObjectDistancia;
            for (int i = 0; i < list3.Count; i++)
            {
                JToken item = list3[i];
                if (c == 1)
                {
                    resudistancia = item.First.ToString();
                    break;
                }

                c = c + 1;
            }

            JObject resultadofinalObjectTiempo = (JObject)JsonConvert.DeserializeObject(strDataTiempoJSON);
            int d = 0;

            IList<JToken> list4 = resultadofinalObjectTiempo;
            for (int i = 0; i < list4.Count; i++)
            {
                JToken item = list4[i];
                if (d == 1)
                {
                    resutiempo = item.First.ToString();
                    break;
                }

                d = d + 1;
            }

            string result = resudistancia + "," + resutiempo;

            return result;
        }
    }
}