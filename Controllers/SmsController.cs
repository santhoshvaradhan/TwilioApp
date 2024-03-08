using IronXL;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Net;
using System.Reflection;
using Twilio;
using Twilio.Rest.Api.V2010.Account;
using System.Net.Http;
using System.Web;
using System.IO;
using Microsoft.IdentityModel.Tokens;
using System.Net.Mail;

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


        [HttpPost("SendSms")]
        public ActionResult UploadFile(IFormFile file ,string year,string doingyear,string dept,string examtype)
        {
            List<string> headers = new List<string>();
            string successMessage = "Sms sent successfully.";
            List<string> values = new List<string>();
            Stream stream = new System.IO.MemoryStream();
            WorkBook workbook = null;



            string fileName = file.FileName;
            file.CopyTo(stream);
            workbook = WorkBook.Load(stream);
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

                if (string.IsNullOrWhiteSpace(sheet.Rows[index].ToString()) == false)
                {
                   
                    if (Convert.ToString(sheet.GetRow(0).RangeAddress) != Convert.ToString(sheet.Rows[index].RangeAddress))
                    {
                       
                        foreach (var cell in sheet.Rows[index])
                        {
                           
                            if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                            {
                               
                                values.Add(cell.ToString());
                            }
                            else
                            {
                                cell.Value = Sms_function();
                                break;
                            }
                        }

                        
                        
                    }
                   
                }
                else
                { 
                    break;
                }
                
            }
             string Sms_function()
            {
                Sms messageObject = FrameMessage(headers, values);


                if (messageObject != null && !string.IsNullOrEmpty(messageObject.Message))
                {
                    Console.WriteLine(messageObject.Message);
                    string accountSid = "ACbb71961314f4d8998bbbba60128b9a65";
                    string authToken = "f4044bf8e17abb2c0117272669d8da01";
                    TwilioClient.Init(accountSid, authToken);

                    var message = MessageResource.Create(
                        body: messageObject.Message,
                        from: new Twilio.Types.PhoneNumber("+14403726082"),
                        to: new Twilio.Types.PhoneNumber("+91" + messageObject.ToNumber)
                    );
                    successMessage = !string.IsNullOrEmpty(message.Sid) ? string.Format("Sms sent successfully to : {0}", messageObject.ToNumber) : "Unable to send sms.";

                    messageObject.ToNumber = string.Empty;
                    messageObject.Message = string.Empty;
                    values.Clear();
                    return successMessage;
                }
                return "Message is empty";
            }
            workbook.SaveAs("D:\\file.xlsx");

            System.IO.File.Move("D:\\file.xlsx", String.Format("D:\\{0}-{1}-{2}-{3}-SMSstatus.xlsx", year, doingyear, dept,examtype));
            Email_function(year,doingyear,dept,examtype);
            return Ok(successMessage);
        }

       

        public static Sms FrameMessage(List<string> keyheader, List<string> vlaues)
        {
           
            Sms messageObject = new Sms();
            for (int i = 0; i < keyheader.Count-1; i++)
            {
                if (keyheader[i] == "SECTION" || keyheader[i] == "ENROLLNO" || keyheader[i] == "NAME")
                {
                    messageObject.Message += String.Format("{0}--{1}\n", keyheader[i], vlaues[i]);
                }
                else if (keyheader[i] == "MOBILENUMBER")
                {
                    messageObject.ToNumber = vlaues[i];

                }
                else if (keyheader[i]=="MESSAGESTATUS")
                {
                    continue;
                }
                else
                {

                    messageObject.Message += String.Format("{0}--{1}\n", keyheader[i], Decisionfunc(vlaues[i]));
                }

            }
            messageObject.Message = "HI THIS IS MESSAGE FROM MAILAM ENGINEERING COLLEGE\nIAT-1 RESULT\n" + messageObject.Message + "Thank you!";          
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
      
        public static void Email_function(string year, string doingyear, string dept,string examtype)
        {
            
            System.Console.WriteLine("Sent");
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("live.smtp.mailtrap.io");
            mail.From = new MailAddress("mailtrap@demomailtrap.com");
            mail.To.Add("sandy.tech02@gmail.com");
            mail.Subject = String.Format("SMS Status of {0}-{1}-{2}-{3}",year,doingyear,dept,examtype);
            mail.Body = "SMS Successully Send to parents,verify message status in below sheet";
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(String.Format("D:\\{0}-{1}-{2}-{3}-SMSstatus.xlsx", year, doingyear, dept,examtype));
            mail.Attachments.Add(attachment);
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("api", "4d90e7d765b6e553a51bcbd8ce692986");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
        }

    }
}
