using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Spire.Xls;
using System.Data;
using System.Net;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace InvoiceAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [AllowAnonymous]
    public class InvoiceController : ControllerBase
    {
        // GET: api/<InvoiceController>
        [HttpGet]
        public OkObjectResult Get()
        {
            Workbook workbook = new Workbook();
            // Get the current directory of the application
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            // Combine the current directory with the relative path to your Excel file
            string filePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "BiometricInputData.xlsx");
            workbook.LoadFromFile(filePath);
            // Assuming you want to read from the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

            // Convert the worksheet data to a DataTable
            DataTable dataTable = worksheet.ExportDataTable();
            //Filter out data table to calculate proper attendance 
            dataTable.Columns["S.No"].ColumnName = "sno";
            dataTable.Columns["NAME"].ColumnName = "name";
            dataTable.Columns["EMP ID"].ColumnName = "empid";
            dataTable.Columns["EMP TYPE"].ColumnName = "emptype";
            dataTable.Columns["MODE"].ColumnName = "mode";
            dataTable.Columns["DATE"].ColumnName = "date";
            dataTable.Columns["TIME"].ColumnName = "time";
            dataTable.Columns["LOCATION"].ColumnName = "location";
            dataTable.Columns["Location Code"].ColumnName = "locationcode";
            dataTable.Columns["Location Name"].ColumnName = "locationname";
            string json = ConvertDataTableToJson(dataTable);
            return Ok(json);
        }
        private string ConvertDataTableToJson(DataTable dataTable)
        {
            // Convert DataTable to JSON using Newtonsoft.Json
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(dataTable);
            return json;
        }

        // GET api/<InvoiceController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<InvoiceController>
        [HttpPost]
        public OkResult Post(IFormFile file)
        {
            var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory() + "/Uploads", 
                DateTime.Now.ToString("yyyy-MM-ddTHH-mm") + file.FileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                file.CopyToAsync(stream);
            }
            return Ok();
        }

        // PUT api/<InvoiceController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<InvoiceController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
