// Title: How to embed an Aspose.Cells column chart as a JPEG image in an HTML email using Content-ID in C#
// Common Searches: C# render Aspose.Cells chart to JPEG and embed in HTML email body | send email with inline chart image using LinkedResource and Content-ID in .NET | Aspose.Cells generate chart image and attach as embedded image in MailMessage | create HTML email with embedded Excel chart image using Aspose.Cells and SmtpClient | use AlternateView to embed JPEG chart from workbook in C# email
// Tags: Aspose.Cells chart rendering to JPEG | embed JPEG image in HTML email C# | LinkedResource Content-ID inline image .NET | AlternateView HTML email with embedded chart | send workbook chart via SmtpClient | memory stream chart image attachment

using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsEmailChart
{
    // The example creates a Workbook, adds sample data, builds a column chart, renders the chart to a JPEG image in a MemoryStream, and constructs a MailMessage with an HTML body that references the image via a Content-ID. The JPEG is added as a LinkedResource to an AlternateView, enabling the image to appear inline when the email is opened. SMTP settings are read from environment variables, and a preview is shown if the configuration is missing.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(1800);

                // 2. Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.SetChartDataRange("A1:B4", true);
                chart.Title.Text = "Quarterly Sales";

                // 3. Render the chart to a JPEG image in memory
                using (MemoryStream imgStream = new MemoryStream())
                {
                    var imgOptions = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Jpeg
                    };
                    chart.ToImage(imgStream, imgOptions);
                    imgStream.Position = 0; // reset for reading

                    // 4. Build the email message
                    MailMessage message = new MailMessage();

                    string fromAddress = Environment.GetEnvironmentVariable("MAIL_FROM") ?? "sender@example.com";
                    string toAddress = Environment.GetEnvironmentVariable("MAIL_TO") ?? "recipient@example.com";

                    message.From = new MailAddress(fromAddress);
                    message.To.Add(new MailAddress(toAddress));
                    message.Subject = "Chart Image Embedded Inline";

                    // 5. Create HTML body referencing the image via Content-ID
                    string contentId = Guid.NewGuid().ToString();
                    string htmlBody = $@"
<html>
<body>
    <h2>Quarterly Sales Chart</h2>
    <p>Below is the chart rendered from an Excel workbook:</p>
    <img src=""cid:{contentId}"" alt=""Chart Image"" />
</body>
</html>";

                    // 6. Create an AlternateView for HTML and attach the image as a LinkedResource
                    AlternateView avHtml = AlternateView.CreateAlternateViewFromString(
                        htmlBody,
                        Encoding.UTF8,
                        MediaTypeNames.Text.Html);

                    LinkedResource chartImage = new LinkedResource(imgStream, MediaTypeNames.Image.Jpeg)
                    {
                        ContentId = contentId,
                        TransferEncoding = TransferEncoding.Base64
                    };

                    avHtml.LinkedResources.Add(chartImage);
                    message.AlternateViews.Add(avHtml);

                    // 7. Configure SMTP client (read from environment variables)
                    string smtpHost = Environment.GetEnvironmentVariable("SMTP_HOST");
                    string smtpPortStr = Environment.GetEnvironmentVariable("SMTP_PORT");
                    string smtpUser = Environment.GetEnvironmentVariable("SMTP_USER");
                    string smtpPass = Environment.GetEnvironmentVariable("SMTP_PASS");
                    bool useSsl = (Environment.GetEnvironmentVariable("SMTP_SSL") ?? "true").ToLower() == "true";

                    if (string.IsNullOrWhiteSpace(smtpHost) || string.IsNullOrWhiteSpace(smtpPortStr))
                    {
                        Console.WriteLine("SMTP configuration not found in environment variables. Email will not be sent.");
                        Console.WriteLine("Set SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS (optional) to enable sending.");
                        Console.WriteLine("Generated email message preview:");
                        Console.WriteLine("-----------------------------------");
                        Console.WriteLine($"From: {fromAddress}");
                        Console.WriteLine($"To:   {toAddress}");
                        Console.WriteLine($"Subject: {message.Subject}");
                        Console.WriteLine("HTML Body:");
                        Console.WriteLine(htmlBody);
                        return;
                    }

                    if (!int.TryParse(smtpPortStr, out int smtpPort))
                    {
                        Console.WriteLine($"Invalid SMTP_PORT value: '{smtpPortStr}'. Email not sent.");
                        return;
                    }

                    using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                    {
                        smtp.EnableSsl = useSsl;

                        if (!string.IsNullOrWhiteSpace(smtpUser))
                        {
                            smtp.Credentials = new NetworkCredential(smtpUser, smtpPass ?? string.Empty);
                        }
                        else
                        {
                            smtp.UseDefaultCredentials = true;
                        }

                        Console.WriteLine("Sending email...");
                        smtp.Send(message);
                        Console.WriteLine("Email sent successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
