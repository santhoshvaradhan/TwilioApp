using IronXL;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Twilio;
using Twilio.Rest.Api.V2010.Account;

namespace TwilioApp.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class SmsController : ControllerBase
    {

        private readonly ILogger<SmsController> _logger;

        public SmsController(ILogger<SmsController> logger)
        {
            _logger = logger;
        }

        [HttpPost(Name = "SendSms")]
        public string Post()
        {
            List<string> headers = new List<string>();
            string successMessage = string.Empty;
            List<string> values = new List<string>();
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "TwilioApp.2023-II-CSE.xlsx";
            WorkBook workbook = null;
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                workbook = WorkBook.Load(stream);
            }
            //  WorkBook workbook = WorkBook.Load(filename: "C:\\Users\\sureshkumar.r\\Downloads\\2023-II-CSE.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            var row1 = sheet[Convert.ToString(sheet.GetRow(0).RangeAddress)];

            foreach (var value in row1)
            {
                if (value != null && !string.IsNullOrEmpty(value?.Value?.ToString()))
                {
                    headers.Add(value.ToString());
                }
            }
            
            for (int index = 0; index < sheet.Rows.Count(); index++)
            {
                if (Convert.ToString(sheet.GetRow(0).RangeAddress) != Convert.ToString(sheet.Rows[index].RangeAddress))
                {
                    foreach (var cell in sheet.Rows[index])
                    {
                        if (cell != null && !string.IsNullOrEmpty(cell?.Value?.ToString()))
                        {
                            values.Add(cell.ToString());
                        }
                    }

                    Sms messageObject = FrameMessage(headers, values);

                    if (messageObject != null && !string.IsNullOrEmpty(messageObject.Message))
                    {
                        string accountSid = "AC671f1a3d410d46e2b2651b5779e29862";
                        string authToken = "edade783c133f59739ecc4ef896d3c61";
                        TwilioClient.Init(accountSid, authToken);

                        var message = MessageResource.Create(
                            body: messageObject.Message,
                            from: new Twilio.Types.PhoneNumber("+13345818542"),
                            to: new Twilio.Types.PhoneNumber("+91" + messageObject.ToNumber)
                        );


                        successMessage = !string.IsNullOrEmpty(message.Sid) ? string.Format("Sms sent successfully to : {0}", messageObject.ToNumber) : "Unable to send sms.";
                        messageObject.ToNumber = string.Empty;
                        messageObject.Message = string.Empty;
                    }
                }

            }

            return successMessage;
        }

        public static Sms FrameMessage(List<string> keyheader, List<string> vlaues)
        {
            string messagebody = string.Empty;
            string mobilenumber = string.Empty;
            Sms messageObject = new Sms();
            for (int i = 0; i < keyheader.Count; i++)
            {
                if (keyheader[i] == "SECTION" || keyheader[i] == "ENROLLNO" || keyheader[i] == "NAME")
                {
                    messagebody += String.Format("{0}--{1}\n", keyheader[i], vlaues[i]);
                }
                else if (keyheader[i] == "MOBILENUMBER")
                {
                    messageObject.ToNumber = vlaues[i];

                }
                else
                {

                    messageObject.Message += String.Format("{0}--{1}\n", keyheader[i], Decisionfunc(vlaues[i]));
                }

            }
            messageObject.Message = "HI THIS IS MESSAGE FROM MAILAM ENGINEERING COLLEGE\nIAT-1 RESULT\n" + messageObject.Message + "Thank you!";
            Console.WriteLine(mobilenumber);
            messagebody = string.Empty;
            mobilenumber = string.Empty;

            return messageObject;
        }

        public static string Decisionfunc(string mark)
        {
            if (mark == "a" || mark == "A" || mark == "absent" || mark == "Absent")
            {
                return "Absent";
            }
            else
            {

                if (!string.IsNullOrEmpty(mark) && Convert.ToInt32(mark) >= 50)
                {
                    return mark + "(Pass)";
                }
                else
                {
                    return mark + "(Fail)";
                }
            }
        }


    }
}