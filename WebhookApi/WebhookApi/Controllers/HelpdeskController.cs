using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Cloud.Dialogflow.V2;
using Google.Protobuf;
using Microsoft.AspNetCore.Mvc;
using Grpc.Core;
using System.Collections;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Internal;
using Microsoft.Extensions.Configuration;
using DotNetCoreApiSample;
using Microsoft.Extensions.Logging;

namespace WebhookApi.Controllers
{
    [Route("helpdesk")]
    public class HelpdeskController : Controller
    {

        private static readonly JsonParser jsonParser = new JsonParser(JsonParser.Settings.Default.WithIgnoreUnknownFields(true));
        private readonly IHostingEnvironment env;
        private readonly ILogger<HelpdeskController> _logger;
        public HelpdeskController(IHostingEnvironment _env, ILogger<HelpdeskController> _log)
        {
            env = _env;
            _logger = _log;
        }

        [HttpPost]
        public async Task<JsonResult> getData()
        {
            try
            {
                _logger.LogInformation("Start Reqst");

                WebhookRequest request;
                using (var reader = new StreamReader(Request.Body))
                {
                    request = jsonParser.Parse<WebhookRequest>(reader);
                }
                _logger.LogInformation(request.QueryResult.ToString());

                var pas = request.QueryResult.Parameters;
                var ListofQuries = pas.Fields.ToList();

                // var VarprinterIssue = ListofQuries[0].Value.ToString().Replace(@"""", "") ?? null;
                //var VarprinterModel = (ListofQuries[1].Value).ToString().Replace(@"""", "") ?? null;
                var askingPrinterIssue = pas.Fields.ContainsKey("printerIssue") && pas.Fields["printerIssue"].ToString().Replace('\"', ' ').Trim().Length > 0;
                var askingPrinterModel = pas.Fields.ContainsKey("PrinterModel") && pas.Fields["PrinterModel"].ToString().Replace('\"', ' ').Trim().Length > 0;
                var response = new WebhookResponse();

                string resolve = "ya user searching our database for printer bt3tk ";
                // string PrinterModel = pas.Fields["PrinterModel"].ToString().Replace('\"', ' ').Trim();
                string printerIssue = pas.Fields["printerIssue"].ToString().Replace('\"', ' ').Trim();
                // string resolve2 = $"ya user searching our database for printer bt3tk {VarprinterModel} with issue{ VarprinterIssue}please wait...";
                //todo search in file
                StringBuilder sb = new StringBuilder();
                //if (!askingPrinterModel)
                //{
                //    sb.Append("could you provide me with printer model???");

                //}
                if (!askingPrinterIssue)
                {
                    sb.Append("what kind of issue are you having?");

                }

                else if (askingPrinterIssue)
                {
                    string filename = "Printer.xlsx";
                    string dataDir = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();
                    string fullPath = Path.Combine(dataDir, filename);
                    var abs = Path.GetFullPath("~/App_Data/Printer.xlsx").Replace("~\\", "");
                    //string x = fullPath;
                    //var byteArrary = System.IO.File.ReadAllBytes(fullPath.ToString());
                    //_logger.LogInformation($"byteArrary  Length {byteArrary.Length}");


                    //string webRootPath = env.WebRootPath;
                    //string contentRootPath = env.ContentRootPath;
                    //var rooted = GetUrlFromAbsolutePath(fullPath);
                    //var GHrequest = HttpContext.Request;
                    //var uriBuilder = new UriBuilder
                    //{
                    //    Host = GHrequest.Host.Host,
                    //    Scheme = GHrequest.Scheme,
                    //    Path = filename
                    //};

                    //if (GHrequest.Host.Port.HasValue)
                    //    uriBuilder.Port = GHrequest.Host.Port.Value;

                    //var url = uriBuilder.ToString();
                    _logger.LogInformation($"workbook XLWorkbook opening from {abs}");
                    try
                    {
                        // Open a SpreadsheetDocument based on a filepath.
                        var workbook = new XLWorkbook(abs);
                        var worksheet = workbook.Worksheet(2);
                        var rows = worksheet.RangeUsed().RowsUsed();
                        _logger.LogInformation($" rows count {rows.ToList().Count}");

                        var cols = worksheet.RangeUsed().RowsUsed();
                        //List<string> issue = new List<string>();
                        //List<string> issueSolve = new List<string>();
                        //foreach (var col in cols)
                        //{
                        //    var colNumber = col.RowNumber();

                        //    string objPage = col.Cell(1).GetString();
                        //    string objElement = col.Cell(2).GetString();
                        //}
                        string result = string.Empty;
                        _logger.LogInformation($" rows loop");
                        foreach (var row in rows)
                        {
                            var rowNumber = row.RowNumber();
                          //  _logger.LogInformation($" rowNumber {rowNumber} rows string {row.Cell(1).GetString()}");

                            var objPage = row.Cell(1).GetString().Split('-').ToArray();
                            string subresult = string.Empty;
                            string objElement = row.Cell(2).GetString();

                            foreach (string text in objPage)
                            {
                                
                                bool contains = ContainsString(text,printerIssue.ToLower(), StringComparison.OrdinalIgnoreCase);
                                _logger.LogInformation($" ======================text {text} do contains {printerIssue} with solve {objElement}");
                                if (contains)
                                {
                                    result = objElement;
                                    _logger.LogInformation($"_______ result {result}");
                                    break;
                                }

                            }
                            //  _logger.LogInformation($" ======================rowNumber {rowNumber} row with issue >  {objPage} rows with solve {objElement}");

                            // bool matchFound = objPage.Any(printerIssue.ToLower().Contains);
                            // var matches = FindMatchs(objPage, printerIssue.ToLower());
                            // _logger.LogInformation($"+++++++========== matchFound {printerIssue.ToLower()}");
                            //   _logger.LogInformation($"+++++++========== matchFound {matches} from issue {printerIssue.ToLower()} ++++++ +====== ");
                            //  if (matchFound)
                            //if (!string.IsNullOrEmpty(matches))
                            //{
                            //    result = objElement;
                            //    _logger.LogInformation($" result {result}");
                            //    break;
                            //}

                        }
                        workbook.Dispose();

                        if (!string.IsNullOrEmpty(result))
                        {
                            sb.Append($"please follow these steps to fix your issue {result}");
                        }
                        else
                        {
                            sb.Append($"couldn't find the issue {printerIssue} for ur printer please call us");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "It broke from open file :(");

                    }

                }
                else
                {

                    if (sb.Length == 0)
                    {
                        sb.Append("Greetings from our Webhook API!");
                    }
                }


                response.FulfillmentText = sb.ToString();

                return Json(response);
            }
            catch (Exception e)
            {
                _logger.LogError(e, "It broke :(");

                return Json(e);

            }

        }
        public static string FindMatchs(string[] array, string filter)
        {
            return array.Where(x => filter.Any(y => x.Contains(y))).FirstOrDefault();
        }
        public static bool ContainsString( string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
        public static string GetUrlFromAbsolutePath(string absolutePath)
        {
            return absolutePath.Replace(Startup.wwwRootFolder, "").Replace(@"\", "/");
        }

    }
}