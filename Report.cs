
using System;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net;
using Syncfusion.XlsIO;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Microsoft.Extensions.Configuration;
namespace Company.Function
{
    //curl -v -u user:password https://<app-name>.azurewebsites.net/api/test/2021-04-25/2022-05-24 --output report.xls
    public static class Report
    {
        static HttpResponseMessage response;
        const string query = "SELECT * FROM c WHERE c.payment_date BETWEEN {from} AND CONCAT({to},' 23:59:60')";
        [FunctionName("Report")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "report/{from}/{to}")] HttpRequestMessage req,
            [CosmosDB("outDatabase", "WebhookCollection", ConnectionStringSetting = "CosmosDbConnectionString", SqlQuery = query)] IEnumerable<object> inputDocument,
            string from,
            string to,
            ILogger log)
        {
           
            var authHeader = req.Headers.Authorization;
            if (authHeader != null && authHeader.ToString().StartsWith("Basic"))
            {
                string encodedUsernamePassword = authHeader.ToString().Substring("Basic ".Length).Trim();

                //the coding should be iso or you could use ASCII and UTF-8 decoder
                Encoding encoding = Encoding.GetEncoding("iso-8859-1");
                string usernamePassword = encoding.GetString(Convert.FromBase64String(encodedUsernamePassword));
                var config = new ConfigurationBuilder()
                            .AddEnvironmentVariables()
                            .Build();
                //string userNameKeyVault = config["WebHookAuth"];
                string userNameKeyVault = Environment.GetEnvironmentVariable("WebHookAuth", EnvironmentVariableTarget.Process);
                log.LogInformation(usernamePassword);
                if (usernamePassword != userNameKeyVault)
                {

                    return new HttpResponseMessage(HttpStatusCode.Unauthorized);
                }

                //Instantiate the spreadsheet creation engine
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //Instantiate the Excel application object
                    IApplication application = excelEngine.Excel;

                    //Assigns default application version
                    application.DefaultVersion = ExcelVersion.Excel2013;

                    //A new workbook is created equivalent to creating a new workbook in Excel
                    //Create a workbook with 1 worksheet
                    IWorkbook workbook = application.Workbooks.Create(1);

                    //Access a worksheet from workbook
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Adding text data
                    worksheet.Range["A1"].Text = "Transaction Date";
                    worksheet.Range["B1"].Text = "DpsTxnRef";
                    worksheet.Range["C1"].Text = "ReCo";
                    worksheet.Range["D1"].Text = "ResponseText";
                    worksheet.Range["E1"].Text = "DpsBillingId";

                    int i = 2;
                    foreach (var e in inputDocument)
                    {
                        //  var jsonToReturn = JsonConvert.SerializeObject(e);
                        var json = JsonConvert.SerializeObject(e);
                        JToken responseBody = JToken.FromObject(e);
                        //var dictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                        string date = responseBody["payment_date"].ToString().Replace("00:00:00 +0000 UTC", "");
                        JToken metadata = responseBody["metadata"];
                        string DpsTxnRef = null;
                        string reCo = null;
                        string responsetext = null;
                        string DpsBillingId = null;


                        if (metadata != null)
                        {
                            DpsTxnRef = metadata["windcave_dpstxnref"] != null ? metadata["windcave_dpstxnref"].ToString() : null;
                            reCo = metadata["windcave_reco"] != null ? metadata["windcave_reco"].ToString() : null;
                            responsetext = metadata["windcave_responsetext"] != null ? metadata["windcave_responsetext"].ToString() : null;
                            DpsBillingId = metadata["windcave_transaction_dpsbillingid"] != null ? metadata["windcave_transaction_dpsbillingid"].ToString() : null;
                        }
                        else
                        {
                            reCo = responseBody["response_code"].ToString();
                            responsetext = responseBody["response_text"].ToString();

                        }

                        worksheet.Range["A" + i].Text = date;
                        worksheet.Range["B" + i].Text = DpsTxnRef;
                        worksheet.Range["C" + i].Text = reCo;
                        worksheet.Range["D" + i].Text = responsetext;
                        worksheet.Range["E" + i].Text = DpsBillingId;
                        i++;
                    }

                    MemoryStream memorystream = new MemoryStream();

                    //Saving the workbook to stream in XLSX format
                    workbook.Version = ExcelVersion.Excel2013;
                    workbook.SaveAs(memorystream);
                    //Create the response to return
                    response = new HttpResponseMessage(HttpStatusCode.OK);

                    //Set the Excel document content response
                    response.Content = new ByteArrayContent(memorystream.ToArray());

                    //Set the contentDisposition as attachment
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = "Report-from-" + from + "-to-" + to + ".xlsx"
                    };
                    //Set the content type as xlsx format mime type
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheet.excel");
                }

                return response;
            }
            else
            {
                return new HttpResponseMessage(HttpStatusCode.Unauthorized);
            }
        }
    }
}
