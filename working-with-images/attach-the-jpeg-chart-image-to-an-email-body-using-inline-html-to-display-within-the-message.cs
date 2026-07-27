using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsEmailDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(7);

            // 2. Add a column chart and bind data
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // 3. Render the chart to a JPEG image in a memory stream
            byte[] jpegBytes;
            using (MemoryStream imgStream = new MemoryStream())
            {
                // Use the ToImage overload that accepts ImageType
                chart.ToImage(imgStream, ImageType.Jpeg);
                jpegBytes = imgStream.ToArray();
            }

            // 4. Convert the JPEG bytes to a Base64 string
            string base64Image = Convert.ToBase64String(jpegBytes);

            // 5. Build the HTML body with an inline image
            string htmlBody = $@"
                <html>
                <body>
                    <h2>Chart Embedded Inline</h2>
                    <img src='data:image/jpeg;base64,{base64Image}' alt='Chart Image' />
                </body>
                </html>";

            // 6. Prepare and send the email (adjust SMTP settings as needed)
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");
            message.To.Add("recipient@example.com");
            message.Subject = "Aspose.Cells Chart as Inline Image";
            message.Body = htmlBody;
            message.IsBodyHtml = true;

            // Example SMTP client configuration – replace with real credentials/host
            SmtpClient smtp = new SmtpClient("smtp.example.com", 587)
            {
                EnableSsl = true,
                Credentials = new NetworkCredential("smtp_user", "smtp_password")
            };

            try
            {
                smtp.Send(message);
                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to send email: " + ex.Message);
            }
        }
    }
}