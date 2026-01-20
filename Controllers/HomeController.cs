using ExcelBirlestirme.Models;
using ExcelBirlestirme.Services;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExcelBirlestirme.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelService _excelService;

        public HomeController(ILogger<HomeController> logger, IExcelService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ProcessFiles(IFormFile mainFile, IFormFile matchFile)
        {
            if (mainFile == null || matchFile == null)
            {
                TempData["Error"] = "Lütfen her iki dosyayı da seçiniz.";
                return RedirectToAction("Index");
            }

            if (!mainFile.FileName.EndsWith(".xlsx") || !matchFile.FileName.EndsWith(".xlsx"))
            {
                TempData["Error"] = "Lütfen geçerli .xlsx dosyaları yükleyiniz.";
                return RedirectToAction("Index");
            }

            try
            {
                using (var mainStream = mainFile.OpenReadStream())
                using (var matchStream = matchFile.OpenReadStream())
                {
                    var result = _excelService.MergeFiles(mainStream, matchStream);

                    if (result.Success)
                    {
                        return File(result.FileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", result.FileName);
                    }
                    else
                    {
                        TempData["Error"] = result.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Bir hata oluştu: " + ex.Message;
            }

            return RedirectToAction("Index");
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
