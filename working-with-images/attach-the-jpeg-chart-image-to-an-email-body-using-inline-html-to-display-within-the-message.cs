using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ChartEmailExample
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
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // 2. Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // 3. Render the chart to a JPEG image in memory
        using (MemoryStream imgStream = new MemoryStream())
        {
            chart.ToImage(imgStream, ImageType.Jpeg); // Use Aspose.Cells Chart.ToImage(Stream, ImageType)

            // 4. Convert the image bytes to a Base64 string
            string base64Image = Convert.ToBase64String(imgStream.ToArray());

            // 5. Build the HTML body with an inline image
            string htmlBody = $@"
                <html>
                <body>
                    <h3>Sales Chart</h3>
                    <img src='data:image/jpeg;base64,{base64Image}' alt='Chart' />
                </body>
                </html>";

            // 6. Prepare the email message
            MailMessage message = new MailMessage();
            message.From = new MailAddress("sender@example.com");          // TODO: replace with actual sender
            message.To.Add("recipient@example.com");                       // TODO: replace with actual recipient
            message.Subject = "Chart Embedded in Email";
            message.Body = htmlBody;
            message.IsBodyHtml = true;

            // 7. Configure SMTP client (replace placeholders with real values)
            SmtpClient smtp = new SmtpClient("smtp.example.com", 587)     // TODO: SMTP server
            {
                Credentials = new NetworkCredential("username", "password"), // TODO: credentials
                EnableSsl = true
            };

            // 8. Send the email
            try
            {
                smtp.Send(message);
                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending email: " + ex.Message);
            }
        }
    }
}