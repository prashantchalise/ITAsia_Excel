using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;

namespace ITAsia.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ExtractData()
        {

            //Create a new DataTable.
            DataTable dt = new DataTable();
            dt.Columns.Add("Funtional Specifications Checklists");
            dt.Columns.Add("Value");

            if (Request != null)
            {

                HttpPostedFileBase file = Request.Files["ExcelFile"];

                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));

                    fileName = fileName.Insert(fileName.IndexOf("."), DateTime.Now.ToString("_ddMMyyhhmmss"));

                    var path = Path.Combine(Server.MapPath("~/ExcelFiles"), fileName);

                    file.SaveAs(path);

                    var wb = new XLWorkbook(path);
                    var ws = wb.Worksheet(1);

                    ViewBag.Path = path;
                    ViewBag.File = fileName;

                    //Loop through the Worksheet rows.

                    IXLCell cell = null;
                    DataRow toInsert = null;
                    bool isright = true;
                    int rowCount = ws.LastRowUsed().RowNumber();

                    for (int index = 3; index <= rowCount; index++)
                    {
                        cell = ws.Cell(index, 1);
                        if (string.IsNullOrEmpty(cell.Value.ToString())) { continue; }
                        else
                        {
                            toInsert = dt.NewRow();
                            
                            if (Regex.IsMatch(cell.Value.ToString(), @"(^\d+\.\d)"))
                            {
                                toInsert[0] = cell.Value.ToString();
                                toInsert[1] = isright ? "✔" : "✘";
                                isright = !isright;
                            }
                            else
                            {
                                toInsert[0] = cell.Value.ToString();
                                toInsert[1] = string.Empty;
                            }

                            dt.Rows.Add(toInsert);
                        }
                    }
                    
                    //workbook manipulation

                    var secondCol = ws.Range("B3:B" + ws.LastRowUsed().RowNumber());
                    var firstCol = ws.Range("A3:A" + ws.LastRowUsed().RowNumber());
                    var totalRows = ws.LastRowUsed().RowNumber();

                    for (int index = 2; index < totalRows; index++)
                    {
                        if (string.IsNullOrEmpty(firstCol.Cell(index, 1).Value.ToString())) { continue; }
                        else
                        {
                            if (Regex.IsMatch(firstCol.Cell(index, 1).Value.ToString(), @"(^\d+\.\d)"))
                            {
                                secondCol.Cell(index, 1).Value = isright ? "✔" : "✘";
                                secondCol.Cell(index, 1).Style.Font.FontColor = isright ? XLColor.Green : XLColor.Red;
                                isright = !isright;
                            }
                            else
                            {
                                secondCol.Cell(index, 1).Value = string.Empty;
                            }
                        }
                    }
 
                    //saving a modified excel file
                    wb.Save();

                }
            }

            return View(dt);
        }

        //download excel file
        public FilePathResult DownloadFile(string path, string file)
        {
            return File(path, "multipart/form-data", file);
        }


    }

}