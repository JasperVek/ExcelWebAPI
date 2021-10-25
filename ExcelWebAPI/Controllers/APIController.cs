using ExcelWebAPI.Domain;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelWebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class APIController : ControllerBase
    {
        private readonly ILogger<APIController> _logger;
        private readonly IWebHostEnvironment _env;

        public APIController(ILogger<APIController> logger, IWebHostEnvironment env)
        {
            _env = env;
            _logger = logger;
        }

        [HttpPost("api/excel")]
        public IModel GetFile(IFormFile uploadedFile)
        {
            if (!CheckIfExcel(uploadedFile))
            {
                return null;
            }
            DemoReader reader = new DemoReader();

            var worksheet = GetWorkSheet(uploadedFile);

            if(worksheet == null)
            {
                return null;
            }
            IModel endObject;
            try
            {
                endObject = reader.ReadModel(new AllInvestments(), worksheet);
            }
            catch (Exception ex)
            {
                _logger.LogError("Error: ", ex);
                throw;
            }           

            if (endObject == null)
            {
                return null;
            }

            // TODO delete excel file
            // TODO convert to JSON output
            return endObject;
        }

        // initialize the Excel worksheet
        private ExcelWorksheet GetWorkSheet(IFormFile uploadedFile)
        {
            // get worksheet
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var dir = _env.ContentRootPath;
            var fullPath = Path.Combine(dir, Path.GetFileName(uploadedFile.FileName));

            using (var fileStream = new FileStream(fullPath, FileMode.Create))
            {
                uploadedFile.CopyTo(fileStream);
            }

            var package = new ExcelPackage(fullPath);
            var worksheet = package.Workbook.Worksheets[0];
            
            if(worksheet == null)
            {
                return null;
            }
            return worksheet;
        }

        // check to see if the file is not null and ends on .xlsx
        private bool CheckIfExcel(IFormFile file)
        {
            if(file != null && file.FileName.EndsWith(".xlsx"))
            {
                return true;
            }
            return false;
        }
    }
}
