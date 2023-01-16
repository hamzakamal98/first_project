using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using projectExcelfile.Models;
using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Data;

namespace projectExcelfile.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var Name = worksheet.Cells[row, 1].Value.ToString();
                        var Email = worksheet.Cells[row, 2].Value.ToString();

                        // Insert the data into the database table
                    }
                }
            }
            return RedirectToAction("Index");
        }
        public async Task<List<User>> Import(IFormFile file)
        {
            var list = new List<User>();
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    var rowcount = worksheet.Dimension.Rows;
                    for (int row = 2; row < rowcount; row++)
                    {
                        list.Add(new User
                        {
                          
                            Name = worksheet.Cells[row, 2].Value.ToString().Trim(),
                            Email = worksheet.Cells[row, 3].Value.ToString().Trim(),
                        });
                    }
                }

            }
            return list;

            var jsonstring = "[{\"Name\"\\ahmed}]";
            var data = JsonConvert.DeserializeObject<List<User>>(jsonstring);
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Email");
            foreach (var item in data)
            {
                dt.Rows.Add(item.Name, item.Email);
            }
        }
   

       

    }
}
