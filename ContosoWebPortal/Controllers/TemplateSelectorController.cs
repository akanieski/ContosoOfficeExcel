using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ContosoWebPortal.Controllers
{
    public class TemplateSelectorController : Controller
    {
        private IGraphService _graphService;
        private IHostingEnvironment _hosting;

        public TemplateSelectorController(IGraphService graphService, IHostingEnvironment hosting)
        {
            _graphService = graphService;
            _hosting = hosting;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("select/{templateId}")]
        public async Task<IActionResult> LoadTemplate(string templateId)
        {
            FileInfo existingFile = new FileInfo(Path.Combine(_hosting.ContentRootPath, "Report.xlsx"));
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Single(w => w.Name == "Cache");
                var r = new Random();
                for (var rowIndex = 1; rowIndex <= 100; rowIndex++)
                {
                    worksheet.InsertRow(2, 1);
                    for (var colIndex = 1; colIndex <= 12; colIndex++)
                    {
                        worksheet.SetValue(2, colIndex, r.NextDouble() * 10000);
                    }
                }
                MemoryStream output = new MemoryStream();
                package.SaveAs(output);
                MemoryStream outToCloud = new MemoryStream(output.ToArray());

                var item = await _graphService.UploadFileToUsersDocuments(this.GetCurrentUserId(), $"document_{DateTime.Now.ToShortDateString().Replace("/", "-")}.xlsx", outToCloud, "contoso/");

                return Ok(item);
            }
        }
    }
}