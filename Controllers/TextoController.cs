using System.Collections.Generic;
using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using OfficeOpenXml;
using System.Collections;
namespace Geoxpress.Controllers
{
    public class TextoController : Controller
    {

        // GET: Texto
        public ActionResult Index()
        {

            return View();
        }

        public ActionResult TExcel() {


            return View();
        }

        public static void Excel(List<Object[]> excel,string nombre,int banco)
        {
            if(excel.Count > 1) {


                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "Geo";
                    excelPackage.Workbook.Properties.Title = "data";
                    excelPackage.Workbook.Properties.Subject = "data";
                    excelPackage.Workbook.Properties.Created = DateTime.Now;

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                    if (banco == 1)
                    {
                        excel.Insert(0,new Object[]{ "Número de Tarjeta", "Nombre del Cliente titular de la tarjeta", "Tipo de vía dirección Básica", "Nombre de Vía dirección Básica", "Numero de dirección Básica", "Otros Datos: INTERIOR DIRECCION BASICA", "Referencia Básica", "Tipo de zona dirección básica", "Nombre de zona de la dirección básica", "Código postal dirección básica", "Descripción del distrito dirección básica", "Descripción de provincia dirección básica", "Descripción de Departamento dirección básica", "Descripción país de dirección básica", "Ubigeo de la dirección Básica", "Tipo de vía dirección Alterna", "Nombre de Vía dirección Alterna", "Numero de dirección Alterna", "Otros Datos: INTERIOR DIRECCION ALTERNA", "Referencia Alterna", "Tipo de zona dirección alterna", "Nombre de zona de la dirección alterna", "Código postal dirección básica", "Descripción del distrito dirección Alterna", "Descripción de provincia dirección Alterna", "Descripción de Departamento dirección Alterna", "Descripción país de dirección Alterna", "Ubigeo de la dirección Alterna", "Prefijo del teléfono 1", "Número de Teléfono 1", "Extensión teléfono 1", "Prefijo teléfono 2", "Número de Teléfono 2", "Extensión teléfono 2", "Prefijo teléfono 3", "Número de Teléfono 3", "Extensión teléfono 3", "Prefijo teléfono 4", "Número de Teléfono 4", "Extensión teléfono 4", "Prefijo teléfono 5", "Número de Teléfono 5", "Extensión teléfono 5", "Tipo de Documento", "Numero de documento", "Código de cliente", "Fecha de Proceso", "Numero de Contrato", "Código de Referencia", "Código de Oficina Destino", "Código de Oficina Gestora", "Número de Incluso Valija", "Indicador de Masivo", "Código de tipo de tarjeta", "Indicador de tipo de solicitud", "Indicador de Tipo de Participe", "INDICADOR VALOR TARJETA", "CODIGO BARRAS", "Otros Datos: Indicador tipo de Seguimiento", "Otros Datos: Indicador Biométrico", "Código Barra número contrato", "Agrupamiento", "Código de segmento del cliente"});
                    }
                    else if (banco == 2)
                    {
                        excel.Insert(0, new Object[] { "INSTITUCION EMISORA", "CODIGO DEL BIN", "TIPO DE EMISION", "CAMPAÑA DE EMISION DE TARJETA", "NUMERO DE CUENTA", "TIPO DE TARJETA", "NUMERO TARJETA", "CODIGO SEGUIMIENTO DE TARJETA", "TIPO DE SERVICIO", "TARJETAHABIENTE - NOMBRE EN EL PLASTICO", "TITULAR DE LA CUENTA - NOMBRE", "TITULAR DE LA CUENTA - APELLIDO PATERNO", "TITULAR DE LA CUENTA - APELLIDO MATERNO", "TIPO DE DOCUMENTO DE IDENTIDAD DEL TITULAR", "NUMERO DEL DOCUMENTO DE IDENTIDAD - TITULAR", "FECHA DE EMISION DE LA TARJETA", "INDICADOR DE ENTREGA DE TARJETA", "DOMICILIO DEL TITULAR - DIRECCION", "DOMICILIO DEL TITULAR - REFERENCIA", "DOMICILIO DEL TITULAR - CODIGO POSTAL", "DOMICILIO DEL TITULAR - CODIGO UBIGEO", "DOMICILIO DEL TITULAR - DEPARTAMENTO", "DOMICILIO DEL TITULAR - PROVINCIA", "DOMICILIO DEL TITULAR - DISTRITO", "DOMICILIO DEL TITULAR - TELEFONO", "TRABAJO DEL TITULAR - DIRECCION", "TRABAJO DEL TITULAR - REFERENCIA", "TRABAJO DEL TITULAR - CODIGO POSTAL", "TRABAJO DEL TITULAR - CODIGO UBIGEO", "TRABAJO DEL TITULAR - DEPARTAMENTO", "TRABAJO DEL TITULAR - PROVINCIA", "TRABAJO DEL TITULAR - DISTRITO", "TRABAJO DEL TITULAR - TELEFONO", "LUGAR DE ENTREGA OPCIONAL - DIRECCION", "LUGAR DE ENTREGA OPCIONAL - REFERENCIA", "LUGAR DE ENTREGA OPCIONAL - CODIGO POSTAL", "LUGAR DE ENTREGA OPCIONAL - CODIGO UBIGEO", "LUGAR DE ENTREGA OPCIONAL - DEPARTAMENTO", "LUGAR DE ENTREGA OPCIONAL - PROVINCIA", "LUGAR DE ENTREGA OPCIONAL - DISTRITO", "LUGAR DE ENTREGA OPCIONAL - TELEFONO", "NOMBRE DE MANDATARIO 1", "TIPO DE DOCUMENTO DE IDENTIDAD DEL MANDATARIO 1", "NUMERO DEL DOCUMENTO DE IDENTIDAD - MANDATARIO 1", "TELEFONO DEL MANDATARIO 1", "NOMBRE DE MANDATARIO 2", "TIPO DE DOCUMENTO DE IDENTIDAD DEL MANDATARIO 2", "NUMERO DEL DOCUMENTO DE IDENTIDAD - MANDATARIO 2", "TELEFONO DEL MANDATARIO 2", "NOMBRE DE EMPRESA", "TIPO DE DESTINO", "CODIGO DE SUCURSAL", "CODIGO DE AGENCIA", "NUMERO DE ACUSE", "CODIGO DE BOLSA DE SEGURIDAD ", "CODIGO DE BOLSA DE TRANSPORTE ", "REQUIERE RECIBO", "INDICADOR DE PACK", "CODIGO ZIP", "FECHA PROCESO EMISOR", "CODIGO RELACIONADOR DE EMPAQUE", "GRUPO DE AFINIDAD", "CANAL DE VENTA", "INDICADOR PARA COURIER", "TIPO DE GESTION DE DISTRIBUCION", "INFORMACION ADICIONAL PARA LA DISTRIBUCION 1", "INFORMACION ADICIONAL PARA LA DISTRIBUCION 2", "INDICADOR ACTIVIDAD", "CODIGO DE PROMOCION", "INDICADOR DE BIOMETRIA", "INDICADOR DE DOCUMENTACION DIGITAL ", "RESERVADO" });
                    }


                    worksheet.Cells["A1"].LoadFromArrays(excel);
                    //Save your file
                    FileInfo fi = new FileInfo(@"C:\Users\SERGIO\Desktop\dbf\excel\"+nombre+".xlsx");
                    excelPackage.SaveAs(fi);
                }
            }
        }
        public static void BBVA(string archivo, string ruta)
        {

            string[] nombre = { ruta + "/files/paperless.txt", ruta + "/files/fuvex.txt", ruta + "/files/diarios.txt", ruta + "/files/tr.txt", ruta + "/files/biometrico.txt", ruta + "/files/fuvexe.txt" };

            string[] paperless = { "0408", "0785" };
            string[] fuvex = { "0781", "0831", "0896" };
            string[] biometrico = { "0504" };

            int[,] separacion =
            {
            {0,20},{20,53},{73,3},{76,50},{126,8},{134,8},{142,50},{192,3},{195,30},{225,5},{230,30},{260,30},{290,30},{320,40},{360,6},{366,3},
            {369,50},{419,8},{427,8},{435,55},{490,3},{493,30},{523,5},{528,30},{558,30},{588,30},{618,40},{658,6},{664,3},{667,7},{674,4},{678,3},{681,7},{688,4},
            {692,3},{695,7},{702,4},{706,3},{709,7},{716,4},{720,3},{723,7},{730,4},{734,1},{735,11},{746,8},{754,10},{764,20},{784,6},{790,4},{794,4},{798,6},
            {804,3},{807,4},{811,1},{812,1},{813,1},{814,12},{826,2},{828,1},{829,27},{856,1},{857,4}
            };

            var Expaperless = new List<object[]>();
            var Exfuvex = new List<object[]>();
            var Exdiarios = new List<object[]>();
            var Extr = new List<object[]>();
            var Exbiometrico = new List<object[]>();
            var Exfuvexe = new List<object[]>();
            string[] dat;


            using (StreamReader leer = new StreamReader(archivo))
            {
                int contador = 1;


                while (!leer.EndOfStream)
                {
                    string linea = leer.ReadLine();

                    if ((contador != 1))
                    {
                        if (linea.Substring(separacion[54, 0], separacion[54, 1]) == "3")
                        {
                            dat = Tseparacion(linea, separacion);
                            Extr.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[3], (contador), "BBVA");

                        }
                        else if (paperless.Contains(linea.Substring(separacion[50, 0], separacion[50, 1])) && linea.Substring(separacion[54, 0], separacion[54, 1]) == "1" &&
                            (linea.Substring(separacion[55, 0], separacion[55, 1]) == "T" || linea.Substring(separacion[54, 0], separacion[54, 1]) == "D") &&
                            (linea.Substring(separacion[61, 0], separacion[55, 1]) == "0" || linea.Substring(separacion[61, 0], separacion[54, 1]) == "1"))
                        {
                            dat = Tseparacion(linea, separacion);
                            Expaperless.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[0], (contador), "BBVA");
                        }

                        else if (fuvex.Contains(linea.Substring(separacion[50, 0], separacion[50, 1])) && linea.Substring(separacion[54, 0], separacion[54, 1]) == "1" &&
                                                    (linea.Substring(separacion[55, 0], separacion[55, 1]) == "T" || linea.Substring(separacion[54, 0], separacion[54, 1]) == "D") &&
                                                    (linea.Substring(separacion[61, 0], separacion[55, 1]) == "0" || linea.Substring(separacion[61, 0], separacion[54, 1]) == "1"))
                        {
                            dat = Tseparacion(linea, separacion);
                            Exfuvex.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[1], (contador), "BBVA");
                        }

                        else if (linea.Substring(separacion[47, 0], 8) == "00110504" && linea.Substring(separacion[54, 0], separacion[54, 1]) == "1" &&
                                 linea.Substring(separacion[55, 0], separacion[55, 1]) == "T")
                        {
                            dat = Tseparacion(linea, separacion);
                            Exfuvexe.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[5], (contador), "BBVA");
                        }

                        else if (biometrico.Contains(linea.Substring(768, 4)) && linea.Substring(separacion[54, 0], separacion[54, 1]) == "1" &&
                                 linea.Substring(separacion[55, 0], separacion[55, 1]) == "T")
                        {
                            dat = Tseparacion(linea, separacion);
                            Exbiometrico.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[4], (contador), "BBVA");
                        }

                        else
                        {
                            dat = Tseparacion(linea, separacion);
                            Exdiarios.Add(new object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62] });
                            crear_texto(linea, nombre[2], (contador), "BBVA");
                        }


                    }

                    contador++;
                }
            }

            Excel(Extr, "tr",1);
            Excel(Expaperless, "papperles",1);
            Excel(Exfuvex, "fuvex",1);
            Excel(Exdiarios, "diarios",1);
            Excel(Exbiometrico, "biometrico",1);
            Excel(Exfuvexe, "fuvexe",1);
        }

        public static void INTERBANK(string archivo, string nombre_base, string ruta)
        {
            

            string[] nombre = {ruta +"/files/"+nombre_base+"_W1.txt",ruta +"/files/" + nombre_base + "_13.txt",ruta +"/files/" + nombre_base + "_TV.txt",ruta +"/files/" + nombre_base + "_TR.txt",
                                ruta +"/files/" + nombre_base + "_otros.txt",ruta +"/files/" + nombre_base + "_GF.txt",ruta +"/files/" + nombre_base + "_RY.txt",ruta +"/files/" + nombre_base + "_01.txt",
                                ruta +"/files/" + nombre_base + "_800080.txt",ruta +"/files/" + nombre_base + "_TJ.txt" };

            int[,] separacion = {
                                    {0,3},{3,6},{9,3},{12,6},{18,16},{34,1},{35,16},{51,17},{68,1},{69,30},{99,30},{129,30},{159,30},{189,1},{190,12},{202,8},{210,1},{211,120},{331,55},{386,10},{396,6},{402,30},
                                    {432,30},{462,30},{492,10},{502,120},{622,55},{677,10},{687,6},{693,30},{723,30},{753,30},{783,10},{793,120},{913,55},{968,10},{978,6},{984,30},{1014,30},{1044,30},{1074,10},
                                    {1084,40},{1124,1},{1125,12},{1137,10},{1147,40},{1187,1},{1188,12},{1200,10},{1210,60},{1270,1},{1271,6},{1277,6},{1283,19},{1302,19},{1321,10},{1331,1},{1332,1},{1333,3},
                                    {1336,8},{1344,17},{1361,4},{1365,3},{1368,1},{1369,1},{1370,70},{1440,40},{1480,1},{1481,3},{1484,1},{1485,1},{1486,519}
                                };

            string[,] tabla ={
                                {"01","ALTAS"},
                                {"02","RENOVACION"},
                                {"03","NUEVA VERSION"},
                                {"04","ALTAS CHIP"},
                                {"05","RENOVACION CHIP"},
                                {"06","NUEVA VERSION CHIP"}
            };

            int contador = 1;

            var ExW1 = new List<object[]>();
            var Ex13 = new List<object[]>();
            var ExTV = new List<object[]>();
            var ExTR = new List<object[]>();
            var ExGF = new List<object[]>();
            var ExRY = new List<object[]>();
            var Ex01 = new List<object[]>();
            var ExOtros = new List<object[]>();
            var Ex800080 = new List<object[]>();
            var ExTJ = new List<object[]>();

            string[] dat;

            using (StreamReader leer = new StreamReader(archivo))
            {


                while (!leer.EndOfStream)
                {

                    string linea = leer.ReadLine();


                    if (linea.Substring(1484, 1) == "S")
                    {
                        dat = Tseparacion(linea, separacion);
                        ExW1.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[0], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6) == "000001" && linea.Substring(9, 3).Contains(tabla[3, 0]))
                    {
                        dat = Tseparacion(linea, separacion);
                        Ex13.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[1], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[3, 0]))
                    {
                        dat = Tseparacion(linea, separacion);
                        ExTV.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[2], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[4, 0]))
                    {
                        dat = Tseparacion(linea, separacion);
                        ExTR.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[3], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[5, 0]) && linea.Substring(3, 6) == "377753")
                    {
                        dat = Tseparacion(linea, separacion);
                        ExRY.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[6], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[5, 0]) && linea.Substring(3, 6) == "456814")
                    {
                        dat = Tseparacion(linea, separacion);
                        Ex01.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[7], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[2, 0]) && linea.Substring(3, 6) == "800080")
                    {
                        dat = Tseparacion(linea, separacion);
                        Ex800080.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[8], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1) == "N" && linea.Substring(12, 6).Trim() == "" && linea.Substring(9, 3).Trim().Contains(tabla[5, 0]))
                    {
                        dat = Tseparacion(linea, separacion);
                        ExGF.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[5], (contador), "INTE");
                    }
                    else if (linea.Substring(1484, 1).Trim() == "" && linea.Substring(3, 3) == "700" && linea.Substring(34, 1) == "3" && linea.Substring(68, 1) == "P")
                    {

                        dat = Tseparacion(linea, separacion);
                        ExTJ.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[9], (contador), "INTE");
                    }
                    else
                    {
                        dat = Tseparacion(linea, separacion);
                        ExOtros.Add(new Object[] { dat[0], dat[1], dat[2], dat[3], dat[4], dat[5], dat[6], dat[7], dat[8], dat[9], dat[10], dat[11], dat[12], dat[13], dat[14], dat[15], dat[16], dat[17], dat[18], dat[19], dat[20], dat[21], dat[22], dat[23], dat[24], dat[25], dat[26], dat[27], dat[28], dat[29], dat[30], dat[31], dat[32], dat[33], dat[34], dat[35], dat[36], dat[37], dat[38], dat[39], dat[40], dat[41], dat[42], dat[43], dat[44], dat[45], dat[46], dat[47], dat[48], dat[49], dat[50], dat[51], dat[52], dat[53], dat[54], dat[55], dat[56], dat[57], dat[58], dat[59], dat[60], dat[61], dat[62], dat[63], dat[64], dat[65], dat[66], dat[67], dat[68], dat[69], dat[70], dat[71] });
                        crear_texto(linea, nombre[4], (contador), "INTE");
                    }


                    contador++;
                }

            }
            Excel(ExW1, "W1",2);
            Excel(Ex13, "13",2);
            Excel(ExTV, "TV",2);
            Excel(ExTR, "TR",2);
            Excel(ExGF, "GF",2);
            Excel(ExRY, "RY",2);
            Excel(Ex01, "01",2);
            Excel(ExOtros, "Otros",2);
            Excel(Ex800080, "800080",2);
            Excel(ExTJ, "TJ",2);
        }

        public static void RIPLEY(string archivo, string nombre_base, string ruta)
        {
            string [] nombre =
            {  
                ruta + "/files/" + nombre_base + "_001.txt",ruta + "/files/" + nombre_base + "_002.txt",ruta + "/files/" + nombre_base + "_003.txt",ruta + "/files/" + nombre_base + "_004.txt"
            };

            using (StreamReader leer = new StreamReader(archivo))
            {


                while (!leer.EndOfStream)
                {
                    string linea = leer.ReadLine();

                    if (linea.Substring(9, 3) == "001")
                    {
                        crear_texto(linea, nombre[0],0, "RIPL");

                    }
                    if (linea.Substring(9, 3) == "002")
                    {
                        crear_texto(linea, nombre[1],0, "RIPL");

                    }
                    if (linea.Substring(9, 3) == "003")
                    {

                        crear_texto(linea, nombre[2],0, "RIPL");
                    }
                    if (linea.Substring(9, 3) == "004")
                    {

                        crear_texto(linea, nombre[3],0, "RIPL");
                    }

                }

            }


        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file, FormCollection formCollection)
        {
            string _banco = formCollection["banco"];

            if (file != null && file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/upload"), fileName);
                var nombre = Path.GetFileNameWithoutExtension(file.FileName);
                var ruta = Server.MapPath("~/upload");
                file.SaveAs(path);

                if(_banco == "1")
                {
                    BBVA(path,ruta);
                }
                if (_banco =="2")
                {
                    INTERBANK(path,nombre,ruta);
                }
                if (_banco == "3")
                {
                    RIPLEY(path,nombre,ruta);
                }
            }

            return RedirectToAction("Index");
        }
        public static string[] Tseparacion(string linea, int[,] separacion)
        {
            int cantidad = separacion.GetLength(0);

            string[] buffer = new string[cantidad];

            for (int i = 0; i < cantidad; i++)
            {

                buffer[i] = linea.Substring(separacion[i, 0], separacion[i, 1]);
            }


            return buffer;
        }

        public static void crear_texto(string linea, string ruta, int contador, string tipo)
        {
            if (tipo == "BBVA")
            {
                if (contador == 2)
                {

                    using (StreamWriter fichero = new StreamWriter(ruta, false))
                    {
                        fichero.WriteLine(linea);
                    }
                }
                else
                {
                    using (StreamWriter fichero = new StreamWriter(ruta, true))
                    {
                        fichero.WriteLine(linea);
                    }

                }
            }
            if (tipo == "INTE")
            {
                if (contador == 1)
                {

                    using (StreamWriter fichero = new StreamWriter(ruta, false))
                    {
                        fichero.WriteLine(linea);
                    }
                }
                else
                {
                    using (StreamWriter fichero = new StreamWriter(ruta, true))
                    {
                        fichero.WriteLine(linea);
                    }

                }

            }
            if (tipo =="RIPL")
            {
                using (StreamWriter fichero = new StreamWriter(ruta, true ))
                {
                    fichero.WriteLine(linea);
                }
            }


        }


    }
}