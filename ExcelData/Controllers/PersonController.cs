using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelData.Models;

namespace ExcelData.Controllers
{
    public class PersonController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file<br>";
                return View("Index");
            }
            else
            {
                if(excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // Read data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<People> listPeople = new List<People>();
                    for(int row = 3; row <= range.Rows.Count; row++)
                    {
                        People p = new Models.People();
                        p.Id = ((Excel.Range)range.Cells[row, 1]).Text;
                        p.FirstName = ((Excel.Range)range.Cells[row, 2]).Text;
                        p.LastName = ((Excel.Range)range.Cells[row, 3]).Text;
                        p.Phone = int.Parse(((Excel.Range)range.Cells[row, 4]).Text);
                        p.City = ((Excel.Range)range.Cells[row, 5]).Text;
                        listPeople.Add(p);
                    }
                    ViewBag.ListPeople = listPeople;
                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrewct<br>";
                    return View("Index");
                }
               
            }
        }
    }
}