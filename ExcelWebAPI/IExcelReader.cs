using ExcelWebAPI.Domain;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelWebAPI
{
    public interface IExcelReader
    {
        IModel ReadModel(IModel model, ExcelWorksheet excelFile);
        IModel FillModel(IModel model, ExcelWorksheet worksheet);
    }
}
