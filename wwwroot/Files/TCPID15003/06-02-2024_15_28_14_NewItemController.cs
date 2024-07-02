using DeliveryOrder.Models;
using DO_Testing.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DO_Testing.Controllers
{
    public class NewItem : Controller
    {
        private readonly EPM01_PR_OfficeAutomationContext _context;
        public NewItem(EPM01_PR_OfficeAutomationContext context)
        {
            _context = context;
        }
        public IActionResult NewItemDOTemplate()
        {
            ViewBag.items = _context.DeliveryOrderTemplateDistinctViews.OrderByDescending(x => x.CreatedDate).ToList();
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        //Items (sheetname)

        public IActionResult ReadExcel()
        {
            var items = ImportExcel<ItemListTemplate>(@"C:\Users\adam.takariyanto\Downloads\DO - Detail Items nyoba.xlsx", "Items");
            return RedirectToAction(nameof(NewItemDOTemplate));
        }

        public List<T> ImportExcel<T>(string excelFilePath, string sheetName)
        {
            List<T> list = new List<T>();
            Type typeOfObject = typeof(T);
            using (IXLWorkbook workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).First();
                var properties = typeOfObject.GetProperties();
                //header column text
                var columns = worksheet.FirstRow().Cells().Select((v, i) => new {Value = v.Value, Index = i+1 }); //index start from 1
                foreach(IXLRow row in worksheet.RowsUsed().Skip(1))
                {
                    T obj = (T)Activator.CreateInstance(typeOfObject);

                    foreach(var prop in properties)
                    {
                        int colIndex = columns.SingleOrDefault(c => c.Value.ToString() == prop.Name.ToString()).Index;
                        var val = row.Cell(colIndex).Value;
                        var type = prop.PropertyType;
                        prop.SetValue(obj, Convert.ChangeType(val, type));
                    }

                    list.Add(obj);
                }
            }

                return list;
        }

        //public ActionResult PostAddMore(AddMoreViewModel model)
        //{
        //    return View();
        //}
    }
}
