using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Mvc;

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
            using (var stream = new FileStream("app_client_secret.json", FileMode.Open, FileAccess.Read))
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


            var range = $"{sheet}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetId, range);
            // Ecexuting Read Operation...
            var response = request.Execute();
            
            // Getting all records from Column...
            IList<IList<object>> values = response.Values;
            
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    
                    foreach(var item in row)
                    {
                        Console.WriteLine(item);
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
       
    }
   
}
