using Aspose.Cells;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Web.Helpers;

namespace Excel.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : ControllerBase
    {

        [HttpPost]
        public async Task<IActionResult> OnPostUploadAsync(IFormFile files, string TableNane)
        {
            string path = Path.Combine(@"C:\Users\Maicol\source\repos\Excel\Excel.API\Excel\" , files.FileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            //Save the uploaded Excel file.
            string fileName = Path.GetFileName(files.FileName);
            string filePath = Path.Combine(path, fileName);
            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                files.CopyTo(stream);
            }
            FileStream test = new FileStream(filePath, FileMode.Open);

            Workbook workbook = new Workbook(test);

            Worksheet worksheet = workbook.Worksheets[0];
            var rows = worksheet.Cells.MaxRow;
            var columns = worksheet.Cells.MaxColumn;
            
            DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, rows, columns, true);
            dataTable.TableName = TableNane;

            test.Close();

            System.IO.File.Delete(filePath);
            var builder = WebApplication.CreateBuilder();
            var conn = builder.Configuration.GetConnectionString("constr");

            using (SqlConnection connection = new SqlConnection(conn))
            {
                connection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    foreach (DataColumn c in dataTable.Columns)
                        bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);

                    bulkCopy.DestinationTableName = dataTable.TableName;
                    try
                    {
                        bulkCopy.WriteToServer(dataTable);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return Ok(Newtonsoft.Json.JsonConvert.SerializeObject(dataTable));
        }
    }
}
