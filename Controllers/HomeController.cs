using Microsoft.AspNetCore.Mvc;
using CrudPruebas.Models;
using System.Diagnostics;

using System.Data;
using System.Data.SqlClient;


using ClosedXML.Excel;

namespace CrudPruebas.Controllers
{
    public class HomeController : Controller
    {
        private readonly string cadenaSQL;

        public HomeController(IConfiguration config)
        {
            cadenaSQL = config.GetConnectionString("cadenaSQL");
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ExportarExcel()
        {
            DataTable tabla_cbcReglaIncosistencia = new DataTable();

            using (var conexion = new SqlConnection(cadenaSQL)) { 
                conexion.Open();
                using (var adapter = new SqlDataAdapter()) {

                    adapter.SelectCommand = new SqlCommand("cbp_reglas_obtener", conexion);
                    adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                   
                   

                    adapter.Fill(tabla_cbcReglaIncosistencia);
                }
            }


            //usar referencias
            //=========== SEGUNDO - INSTALAR ClosedXML ===========
            using (var libro = new XLWorkbook()) {

                tabla_cbcReglaIncosistencia.TableName = "cbc_regla_inconsistencia";
                var hoja = libro.Worksheets.Add(tabla_cbcReglaIncosistencia);
                hoja.ColumnsUsed().AdjustToContents();

                using (var memoria = new MemoryStream()) {

                    libro.SaveAs(memoria);

                    var nombreExcel = string.Concat("Reporte ", DateTime.Now.ToString(), ".xlsx");

                    return File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                }
            }



        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}