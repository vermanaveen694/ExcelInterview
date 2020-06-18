using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using InterviewExcel.Models;
using Microsoft.AspNetCore.Mvc;
namespace InterviewExcel.Controllers
{
    [Produces("application/json")]
    [Route("api/User")]
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [Route("UpdateRecord")]
        public bool UpdateRecord([FromBody]User user1)
        {
            // string url = "C:\\Users\\w.8\\Desktop\\InterViewTest\\InterviewExcel\\editableForm.xlsx";
           //Path Get Excel File
            string url = Environment.CurrentDirectory + @"\\editableForm.xlsx";
            // Open the existing excel file
            var workbook = new XLWorkbook(url); 
            var worksheet = workbook.Worksheets.Worksheet(1);
            //Update Value in columns In Excel
            worksheet.Cell("D5").SetValue(user1.Name  + " "+ user1.Address);
            worksheet.Cell("D6").SetValue(user1.PAN );
            worksheet.Cell("D7").SetValue(user1.FinacialYear);
            worksheet.Cell("A47").SetValue(user1.Place);
            worksheet.Cell("A48").SetValue(user1.Date);
            worksheet.Cell("A49").SetValue(user1.Designation);
            workbook.Save();
            return true;
        }




    }
}