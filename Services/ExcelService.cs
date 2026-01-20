using ExcelBirlestirme.Models;
using OfficeOpenXml;

namespace ExcelBirlestirme.Services
{
    public class ExcelService : IExcelService
    {
        public ProcessResult MergeFiles(Stream mainFileStream, Stream matchFileStream)
        {
            try
            {
                // Note: LicenseContext is set in Program.cs globally

                using (var mainPackage = new ExcelPackage(mainFileStream))
                using (var matchPackage = new ExcelPackage(matchFileStream))
                {
                    // 1. Validate Worksheets
                    if (mainPackage.Workbook.Worksheets.Count == 0)
                        return new ProcessResult { Success = false, Message = "Ana dosyada çalışma sayfası bulunamadı." };
                    if (matchPackage.Workbook.Worksheets.Count == 0)
                        return new ProcessResult { Success = false, Message = "Eşleştirme dosyasında çalışma sayfası bulunamadı." };

                    var mainSheet = mainPackage.Workbook.Worksheets[0];
                    var matchSheet = matchPackage.Workbook.Worksheets[0];

                    // 2. Read Match Data (Column A from Match File)
                    var matchValues = new HashSet<string>();
                    int matchRowCount = matchSheet.Dimension?.Rows ?? 0;

                    for (int row = 2; row <= matchRowCount; row++)
                    {
                        var val = matchSheet.Cells[row, 1].Text?.Trim(); // Column A
                        if (!string.IsNullOrEmpty(val))
                        {
                            matchValues.Add(val);
                        }
                    }

                    // 3. Process Main File (Check Column G against Match Data)
                    var resultRows = new List<List<string>>();
                    int mainRowCount = mainSheet.Dimension?.Rows ?? 0;
                    int matchesFound = 0;

                    // Header row
                    var headerRow = new List<string>();
                    for (int col = 1; col <= 12; col++)
                    {
                        headerRow.Add(mainSheet.Cells[1, col].Text);
                    }
                    resultRows.Add(headerRow);

                    for (int row = 2; row <= mainRowCount; row++)
                    {
                        var checkVal = mainSheet.Cells[row, 7].Text?.Trim(); // Column G is index 7

                        if (!string.IsNullOrEmpty(checkVal) && matchValues.Contains(checkVal))
                        {
                            var rowData = new List<string>();
                            for (int col = 1; col <= 12; col++) // Take first 12 columns
                            {
                                rowData.Add(mainSheet.Cells[row, col].Text);
                            }
                            resultRows.Add(rowData);
                            matchesFound++;
                        }
                    }

                    // 4. Generate Output File
                    using (var resultPackage = new ExcelPackage())
                    {
                        var ws = resultPackage.Workbook.Worksheets.Add("Sonuclar");

                        for (int r = 0; r < resultRows.Count; r++)
                        {
                            for (int c = 0; c < resultRows[r].Count; c++)
                            {
                                ws.Cells[r + 1, c + 1].Value = resultRows[r][c];
                            }
                        }

                        // Auto fit columns
                        if (ws.Dimension != null)
                        {
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        }

                        return new ProcessResult
                        {
                            Success = true,
                            FileContent = resultPackage.GetAsByteArray(),
                            FileName = "Birlestirilmis_Sonuc.xlsx",
                            MatchCount = matchesFound,
                            Message = $"İşlem başarılı. {matchesFound} adet eşleşme bulundu."
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                return new ProcessResult { Success = false, Message = $"Hata oluştu: {ex.Message}" };
            }
        }
    }
}
