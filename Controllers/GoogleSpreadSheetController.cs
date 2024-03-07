using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Mvc;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;

namespace TwilioApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GoogleSpreadSheetController : Controller
    {
        public string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public string ApplicationName = "GooglesheetAPI";
        public string sheet = "IAT-1";
        public string SpreadsheetId = "1BuADmesS50Kr_d7hUnYmR1tXU08boqt2K-Ux2bnYM7w";
        SheetsService service;
        [HttpGet]
        public void Index()
        {
            GoogleCredential credential;
            //Reading Credentials File...
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            // Creating Google Sheets API service...
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            var range = $"{sheet}!A:E";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);
            // Ecexuting Read Operation...
            var response = request.Execute();
            // Getting all records from Column A to E...
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Writing Data on Console...
                    Console.WriteLine("{0} | {1} | {2} | {3} | {4} ", row[0], row[1], row[2], row[3], row[4]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
       
    }
   
}
