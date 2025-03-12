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
            headers.Add("Best");
            headers.Add("Average");
            headers.Add("STD DEV");
            newDocument.Add(headers);

            using (var stream = new MemoryStream()) {
                request.ExcelFile.CopyTo(stream);
                
                using (var workbook = new XLWorkbook(stream)) {
                    
                    var worksheet = workbook.Worksheets.First();
                    
                    var rows = worksheet.Rows().Skip(1);
                    var rowIndex = 1;
                    var lastBestResultRowNumber = request.LastBestResultRowNumber;
                    var bestResults = new List<double>();
                    var testNumber = 0;
                    foreach (var row in rows) {
                        rowIndex++;

                        if (rowIndex == lastBestResultRowNumber)
                        {
                            lastBestResultRowNumber += 20;
                            testNumber++;
                            var currentRow = new List<string>();
                            currentRow.Add(row.Cell(1).Value.ToString());//Benchmark
                            currentRow.Add(row.Cell(5).Value.ToString());//Dimensions
                            currentRow.Add(row.Cell(3).Value.ToString());//Iterations
                            currentRow.Add(row.Cell(12).Value.ToString());//Best Results
                            bestResults.Add(row.Cell(12).Value.GetNumber());
                            if (testNumber == 30) {

                                double average = bestResults.Average();
                                currentRow.Add(average.ToString());

                                double desviacionEstandar = Math.Sqrt(bestResults.Select(v => Math.Pow(v - average, 2)).Average());

                                currentRow.Add(desviacionEstandar.ToString());
                            }
                            newDocument.Add(currentRow);
                        }
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
