using ClosedXML.Excel;
using HMResultsReader.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text;

namespace HMResultsReader.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelFileReaderController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<ExcelFileReaderController> _logger;

        public ExcelFileReaderController(ILogger<ExcelFileReaderController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpPost("ProcessFile")]
        [ProducesResponseType(typeof(string), 200)]
        public IActionResult ProcessFile(DTORequest request) {
            var newDocument = new List<List<string>>();

            var headers = new List<string>();
            headers.Add("Benchmark");
            headers.Add("Dimensions");
            headers.Add("Iterations");
            headers.Add("Harmony Memory Iteration");
            headers.Add("Min Search Range");
            headers.Add("Max Search Range");
            headers.Add("Adjustment Rate");
            headers.Add("Harmony Memory Size");
            headers.Add("Acceptance Rate");
            headers.Add("Best Result Static");
            headers.Add("Best Result Static Min");
            headers.Add("Best Result Static Mean");
            newDocument.Add(headers);

            using (var stream = new MemoryStream()) {
                request.ExcelFile.CopyTo(stream);
                
                using (var workbook = new XLWorkbook(stream)) {
                    
                    var worksheet = workbook.Worksheets.First();
                    
                    var rows = worksheet.Rows().Skip(1);
                    foreach (var row in rows) {
                        var currentRow = new List<string>();
                        currentRow.Add(row.Cell(1).Value.ToString());
                        currentRow.Add(row.Cell(2).Value.ToString());
                        currentRow.Add(row.Cell(3).Value.ToString());
                        currentRow.Add(row.Cell(4).Value.ToString());
                        currentRow.Add(row.Cell(5).Value.ToString());
                        currentRow.Add(row.Cell(6).Value.ToString());
                        currentRow.Add(row.Cell(7).Value.ToString());
                        currentRow.Add(row.Cell(8).Value.ToString());
                        currentRow.Add(row.Cell(9).Value.ToString());
                        currentRow.Add(row.Cell(10).Value.ToString());
                        currentRow.Add(row.Cell(11).Value.ToString());
                        currentRow.Add(row.Cell(12).Value.ToString());
                        newDocument.Add(currentRow);
                    }
                }
            }
            
            var strBuilder = new StringBuilder();

            for (var i = 0; i < newDocument.Count(); i++) 
            {
                var newDocumentRow = newDocument.ElementAt(i);
                
                string csv = String.Join(",", newDocumentRow);
                strBuilder.AppendLine(csv);
            }

            var csvBytes = Encoding.UTF8.GetBytes(strBuilder.ToString());
            return File(csvBytes, "text/csv", "export.csv");
        }
    }
}
