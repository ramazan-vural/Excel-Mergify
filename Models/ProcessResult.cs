namespace ExcelBirlestirme.Models
{
    public class ProcessResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public byte[] FileContent { get; set; }
        public string FileName { get; set; }
        public int MatchCount { get; set; }
    }
}
