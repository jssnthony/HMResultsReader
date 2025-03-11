using ClosedXML.Excel;
using HMResultsReader.Models;
using Microsoft.AspNetCore.Mvc;

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
        public DTOResponse ProcessFile(DTORequest request) {
            var response = new DTOResponse();
            using (var stream = new MemoryStream()) {
                request.ExcelFile.CopyTo(stream);
                
                using (var workbook = new XLWorkbook(stream)) { 


                }
            }

            return response;
        }
    }
}
