using System.Net.Http.Json;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text;
using System.Web;

namespace SendMessageFromWhatsApp
{
    internal class Program
    {
        static async Task SendSms(string apiUrl, string apiKey, string sender, string recipient, string message)
        {
           
        }
        static async Task Main(string[] args)
        {
          var  credentials = new Dictionary<string, string>
            {
                { "username", "DEXEFCOAPI" }, // Replace with your Victory Link username
                { "password", "L/Bsy3s]+&" }, // Replace with your Victory Link password
                { "language", "English" }, // Replace with the language code
                { "sender", "DexefCompany" }   // Replace with your sender ID
            };
          var data = new Dictionary<string, string>
          {
              { "message", "Hello World" },    // Replace with your message content
              { "to", "+201030454205" }        // Replace with the recipient phone number
          };
          var parameters = new Dictionary<string, string>
          {
              { "UserName", credentials["username"] },
              { "Password", credentials["password"] },
              { "SMSText", data["message"] },
              { "SMSLang", credentials["language"] },
              { "SMSSender", credentials["sender"] },
              { "SMSReceiver", data["to"] }
          };


            string apiUrl = "https://smsvas.vlserv.com/VLSMSPlatformResellerAPI/NewSendingAPI/api/SMSSender/SendToMany"; 

            string apiKey = "L/Bsy3s]+&"; // Replace with your Victory Link API key
            string sender = "DEXEFCOAPI"; // Replace with your sender ID
            string recipient = "+201030454205"; // Replace with the recipient phone number
            string message = "888"; // Re
             using (HttpClient client = new HttpClient())
             {


                //var payload = new
                // {
                //     apiKey = apiKey,
                //     sender = sender,
                //     recipient = recipient,
                //     message = message
                // };
                // //pnBdQQg5SDizttpmktguwA== 
                 var jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(parameters);
                 var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                 try
                 {
                     HttpResponseMessage response = await client.PostAsync(apiUrl, content);
                     response.EnsureSuccessStatusCode();
                     string responseBody = await response.Content.ReadAsStringAsync();
                     Console.WriteLine("SMS sent successfully. Response: " + responseBody);
                 }
                 catch (HttpRequestException e)
                 {
                     Console.WriteLine("Error sending SMS: " + e.Message);
                 }
             }
            //            var contacts = LoadContactsFromExcel("C:\\Users\\Ahmed E. Rizk\\Desktop\\contacts.xlsx");
            //IWebDriver driver = new ChromeDriver();

            //try
            //{
            //    // افتح واتساب ويب
            //    driver.Navigate().GoToUrl("https://web.whatsapp.com");

            //    // انتظر المستخدم لمسح رمز QR
            //    Console.WriteLine("Please scan the QR code with your phone.");
            //    Thread.Sleep(15000); // انتظر 15 ثانية لمسح رمز QR

            //    foreach (var contact in contacts)
            //    {
            //        SendMessage(driver, contact.Item1, contact.Item2);
            //        Thread.Sleep(3000); // انتظر 3 ثوانٍ بين كل رسالة
            //    }
            //}
            //finally
            //{
            //    // أغلق المتصفح
            //    driver.Quit();
            //}
        }




        static void SendMessage(IWebDriver driver, string number, string message)
        {
            // تكوين رابط wa.me للرقم
            string url = $"https://wa.me/{number}?text={HttpUtility.UrlEncode(message)}";

            // افتح الرابط
            driver.Navigate().GoToUrl(url);
            Thread.Sleep(5000); // انتظر 5 ثوانٍ للتأكد من تحميل الصفحة

            // اضغط على زر "استمرار إلى الدردشة" إذا كان موجودًا
            try
            {
                var continueButton = driver.FindElement(By.XPath("//span[contains(text(),'Continue to Chat')]"));
                continueButton.Click();
                Thread.Sleep(5000);
                var useWebButton = driver.FindElement(By.XPath("//span[contains(text(),'use WhatsApp Web') and contains(@class, '_advp') and contains(@class, '_aeam')]"));

                useWebButton.Click();
                Thread.Sleep(5000);
                // انتظر 5 ثوانٍ لتحميل صفحة الدردشة
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"Failed to find the continue button for {number}. Element not found.");
            }

            // اضغط على زر الإرسال إذا كان موجودًا
            try
            {
                var sendButton = driver.FindElement(By.XPath("//button[@data-tab='11' and @aria-label='Send']"));
              sendButton.Click();
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"Failed to send message to {number}. Element not found.");
            }
        }


        static List<(string, string)> LoadContactsFromExcel(string filePath)
        {
            var contacts = new List<(string, string)>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // تخطي الصف الأول الذي يحتوي على رؤوس الأعمدة

                foreach (var row in rows)
                {
                    var number = row.Cell(1).GetString();
                    var message = row.Cell(2).GetString();
                    contacts.Add((number, message));
                }
            }

            return contacts;
        }
    }
}
