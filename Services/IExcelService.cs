using ExcelBirlestirme.Models;

namespace ExcelBirlestirme.Services
{
    public interface IExcelService
    {
        ProcessResult MergeFiles(Stream mainFileStream, Stream matchFileStream);
    }
}
