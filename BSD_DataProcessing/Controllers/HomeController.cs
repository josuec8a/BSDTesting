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

namespace BSD_DataProcessing.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _environment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> UploadExcel(Microsoft.AspNetCore.Http.IFormFile fileupload)
        {
            var dt = new System.Data.DataTable();
            //Checking file content length and Extension must be .xlsx  
            if (fileupload != null)
            {
                if (fileupload.Length > 0 && fileupload.ContentType == "application/vnd.ms-excel" || fileupload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string webRootPath = _environment.WebRootPath + "\\Upload\\";
                    string webRootDownloadPath = _environment.WebRootPath + "\\Upload\\" + "\\Download\\";

                    if (!Directory.Exists(webRootPath))
                    {
                        Directory.CreateDirectory(webRootPath);
                    }

                    string id = Guid.NewGuid().ToString();
                    string fileName = $"{Guid.NewGuid().ToString()}{ Path.GetExtension(fileupload.FileName)}";
                    string filePath = $"{webRootPath}{fileName}";

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
                                    dt.Columns.Add(cell.Value.ToString());
                                }
                                FirstRow = false;
                            }
                            else
                            {
                                if (docIds == null)
                                    docIds = new List<string>();

                                //Adding a Row in datatable  
                                dt.Rows.Add();
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

                    List<string> retryList = null;
                    int attachNumber = 1;

                    DataTable dtDocument = await ProcessData(webRootDownloadPath, docIds, retryList, Constants.ApiUrl, attachNumber);

                    if (dtDocument != null)
                    {
                        if (retryList != null)
                        {
                            var errorDocs = retryList;
                            dtDocument.Merge(await ProcessData(webRootDownloadPath, errorDocs, retryList, Constants.ApiUrl, 2));
                        }

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

                            return fileContentResult;
                        }
                    }
                }
                else
                {
                    //If file extension of the uploaded file is different then .xlsx  
                    ViewBag.Message = "Please select file with .xlsx extension!";
                }
            }
            return RedirectToAction("Index");//View(dt);
        }

        public async Task LogActivity(string path, string textline)
        {
            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter($"{path}Processed.txt", true))
            {
                await file.WriteLineAsync(textline);
            }
        }

        private async Task<DataTable> ProcessData(string webRootDownloadPath, List<string> docIds, List<string> retryList, string apiUrl, int attachNumber)
        {
            DataTable dtDocument = null;

            foreach (string docId in docIds)
            {
                if (dtDocument == null)
                {
                    dtDocument = new DataTable("DB");
                    dtDocument.Columns.Add("DocId", typeof(string));
                    //add columns
                    Constants.GetFields.ForEach(e =>
                    {
                        dtDocument.Columns.Add(e.Name, typeof(string));
                    });
                }

                try
                {
                    var response = await DoHttp(apiUrl, docId, attachNumber);

                    if (response.RawBytes == null)
                    {
                        if (retryList == null)
                            retryList = new List<string>();

                        retryList.Add(docId);

                        await LogActivity(webRootDownloadPath, $"{docId}, accion: reprocesar attachment 2");
                    }
                    else
                    {
                        await LogActivity(webRootDownloadPath, docId);

                        if (!Directory.Exists(webRootDownloadPath))
                        {
                            Directory.CreateDirectory(webRootDownloadPath);
                        }

                        string newFilePath = $"{webRootDownloadPath}{docId}.xlsx";
                        System.IO.File.WriteAllBytes(newFilePath, response.RawBytes);

                        using (XLWorkbook workbook = new XLWorkbook(newFilePath))
                        {
                            IXLWorksheet worksheet = workbook.Worksheet(1);

                            //mapping fields
                            var row = dtDocument.NewRow();
                            var _docId = $"{docId}{(attachNumber > 1 ? " - " + attachNumber.ToString() : string.Empty)}";
                            row["DocId"] = _docId;

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
                        }
                    }
                }
                catch (Exception ex)
                {
                    //var row = dtDocument.NewRow();
                    //row["DocId"] = $"{docId}, error";
                    ////row["DocId"] = $"{docId}, error: {ex.Message}";

                    //dtDocument.Rows.Add(row);

                    await LogActivity(webRootDownloadPath, $"{docId}, error: {ex.Message}");

                }
            }

            return dtDocument;
        }

        public async Task<IRestResponse> DoHttp(string apiUrl, string docId, int attachNumber = 1)
        {
            var request = new RestRequest(Method.GET);
            //var client = new RestClient("https://localhost:44332/file")
            var client = new RestClient($"{apiUrl}?documentId={docId}&attachmentNumber={attachNumber}&contentType=excel12book")
            {
                Timeout = 800
            };

            return await client.ExecuteAsync(request);
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
