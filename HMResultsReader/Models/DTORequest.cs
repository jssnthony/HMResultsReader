namespace HMResultsReader.Models
{
    public class DTORequest
    {
        public IFormFile ExcelFile { get; set; }
        public int LastBestResultRowNumber { get; set; }
    }
}
