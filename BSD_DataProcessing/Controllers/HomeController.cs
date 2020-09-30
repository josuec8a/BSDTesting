using System.Net.Http;
using ClosedXML.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using BSD_DataProcessing.Models;
using Microsoft.AspNetCore.Hosting;
using System.Reflection;
using System.Net;
using RestSharp;
using System.Data;
using BSD_DataProcessing.Helpers;

namespace BSD_DataProcessing.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly string _webRootDownloadPath;
        private readonly string _webRootPath;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
            _webRootDownloadPath = _environment.WebRootPath + "\\Upload\\" + "\\Download\\";
            _webRootPath = _environment.WebRootPath + "\\Upload\\";
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> UploadExcel(Microsoft.AspNetCore.Http.IFormFile fileupload)
        {
            var dtDocumentIdList = new System.Data.DataTable();
            //Checking file content length and Extension must be .xlsx  
            if (fileupload != null)
            {
                if (fileupload.Length > 0 && fileupload.ContentType == "application/vnd.ms-excel" || fileupload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {

                    if (!Directory.Exists(_webRootPath))
                        Directory.CreateDirectory(_webRootPath);

                    string id = Guid.NewGuid().ToString();
                    string fileName = $"{Guid.NewGuid().ToString()}{ Path.GetExtension(fileupload.FileName)}";
                    string filePath = $"{_webRootPath}{fileName}";

                    using (FileStream fs = System.IO.File.Create(filePath))
                    {
                        await fileupload.CopyToAsync(fs);
                        await fs.FlushAsync();
                    }

                    List<string> docIds = null;
                    //Started reading the Excel file.  
                    using (XLWorkbook workbook = new XLWorkbook(filePath))
                    {
                        IXLWorksheet worksheet = workbook.Worksheet(1);
                        bool FirstRow = true;
                        //Range for reading the cells based on the last cell used.  
                        string readRange = "1:1";
                        foreach (IXLRow row in worksheet.RowsUsed())
                        {
                            //If Reading the First Row (used) then add them as column name  
                            if (FirstRow)
                            {
                                //Checking the Last cellused for column generation in datatable  
                                readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
                                foreach (IXLCell cell in row.Cells(readRange))
                                {
                                    dtDocumentIdList.Columns.Add(cell.Value.ToString());
                                }
                                FirstRow = false;
                            }
                            else
                            {
                                if (docIds == null)
                                    docIds = new List<string>();

                                //Adding a Row in datatable  
                                dtDocumentIdList.Rows.Add();
                                int cellIndex = 0;
                                //Updating the values of datatable  
                                foreach (IXLCell cell in row.Cells(readRange))
                                {
                                    docIds.Add(cell.Value.ToString());
                                    //dt.Rows[dt.Rows.Count - 1][cellIndex] = cell.Value.ToString();
                                    cellIndex++;
                                }
                            }
                        }
                        //If no data in Excel file  
                        if (FirstRow)
                        {
                            ViewBag.Message = "Empty Excel File!";
                        }
                    }

                    if (docIds == null) return NotFound();


                    if (!Directory.Exists(_webRootDownloadPath))
                        Directory.CreateDirectory(_webRootDownloadPath);

                    //descarga adjunto 1
                    var taskResult1 = await Task.WhenAll(docIds.Select(doc => DownloadFileAsync(doc, 1)));

                    await LogActivity(_webRootDownloadPath, $"taskResult1.Lenght: {taskResult1.Length}, OK");

                    //descarga adjunto 2
                    var taskResult2 = await Task.WhenAll(taskResult1.Select(doc => DownloadFileAsync(doc, 2)));

                    await LogActivity(_webRootDownloadPath, $"taskResult2.Lenght: {taskResult2.Length}, Error");

                    //lista a procesar
                    List<string> processLst = docIds.Where(p => !taskResult2.Contains(p)).ToList();

                    var dtDocument = new DataTable("DB");
                    dtDocument.Columns.Add("DocId", typeof(string));
                    //add columns
                    Constants.GetFields.ForEach(e =>
                    {
                        dtDocument.Columns.Add(e.Name, typeof(string));
                    });

                    await Task.WhenAll(processLst.Select(s => ProcessDocument(dtDocument, s)));

                    if (dtDocument.Rows.Count > 0)
                    {
                        using (XLWorkbook workbook = new XLWorkbook(filePath))
                        {
                            IXLWorksheet dbSheet = workbook.AddWorksheet(dtDocument, "DB");
                            workbook.Save();

                            var bytes = await System.IO.File.ReadAllBytesAsync(filePath);

                            const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                            HttpContext.Response.ContentType = contentType;
                            HttpContext.Response.Headers.Add("Access-Control-Expose-Headers", "Content-Disposition");

                            var fileContentResult = new FileContentResult(bytes, contentType)
                            {
                                FileDownloadName = fileName
                            };

                            TempData["Ok"] = true;

                            return fileContentResult;
                        }
                    }
                    else
                        TempData["Ok"] = false;
                }
            }
            else
            {
                TempData["SelectFile"] = true;
            }
            return RedirectToAction("Index");
        }

        public async Task LogActivity(string path, string textline)
        {
            using (StreamWriter file =
                new StreamWriter($"{path}Processed.txt", true))
            {
                await file.WriteLineAsync(textline);
            }
        }

        public async Task<string> DownloadFileAsync(string docId, int attachNumber = 1)
        {
            string ret = string.Empty;
            try
            {
                string url = $"{Constants.ApiUrl}?documentId={docId}&attachmentNumber={attachNumber}&contentType=excel12book";

                string filePath = $"{_webRootDownloadPath}{docId}.xlsx";

                using (var client = new MyWebClient(3000))
                {
                    client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

                    await client.DownloadFileTaskAsync(new Uri(url), filePath);
                }
            }
            catch
            {
                ret = docId;
            }

            return ret;
        }

        private async Task ProcessDocument(DataTable dtDocument, string docId)
        {
            try
            {
                string newFilePath = $"{_webRootDownloadPath}{docId}.xlsx";

                using (XLWorkbook workbook = new XLWorkbook(newFilePath))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);

                    //mapping fields
                    var row = dtDocument.NewRow();
                    row["DocId"] = docId;

                    foreach (Fields f in Constants.GetFields)
                    {
                        object value = worksheet.Cell(f.CellPosition).Value;
                        if (f.FormatType == "%")
                            row[f.Name] = $"{value.ToString()} %";
                        if (f.FormatType == "text")
                            row[f.Name] = value.ToString();
                        if (f.FormatType == "number")
                        {
                            decimal.TryParse(value.ToString(), out decimal outDec);
                            row[f.Name] = outDec;
                        }
                    }
                    dtDocument.Rows.Add(row);
                    await LogActivity(_webRootDownloadPath, $"{docId}, OK");

                }
            }
            catch (Exception ex)
            {
                var row = dtDocument.NewRow();
                row["DocId"] = $"{docId}, error: {ex.Message}";
                dtDocument.Rows.Add(row);

                await LogActivity(_webRootDownloadPath, $"{docId}, error: {ex.Message}");
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
