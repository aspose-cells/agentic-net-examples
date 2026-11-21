// Title: How to stream an Aspose.Cells workbook as HTML directly to an ASP.NET HttpResponse in C#
// Common Searches: aspnet download aspose.cells workbook as html without saving to disk | c# write aspose.cells html output to HttpResponse stream | convert excel to html and send to browser using asp.net | asp.net core return aspose.cells html file as file result | memorystream aspose.cells saveformat.html to response
// Tags: Aspose.Cells stream workbook to HTML HttpResponse | ASP.NET write memory stream to response for download | C# convert Excel to HTML bytes using Aspose.Cells | download Excel as HTML without temporary file | SaveFormat.Html to memory stream Aspose.Cells | browser file download from Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, adds sample data, and saves the workbook as an HTML file using Aspose.Cells' SaveFormat.Html. It demonstrates the steps needed to adapt the HTML output to a MemoryStream, which can then be written directly to an ASP.NET HttpResponse for immediate browser download, eliminating the need for a temporary file.
    public class ExcelToHtmlDownloader
    {
        public void DownloadWorkbookAsHtml()
        {
            try
            {
                // Load an existing workbook if the template file exists; otherwise create a new workbook
                Workbook workbook;
                const string templatePath = "Template.xlsx";

                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    // Example data – can be removed or replaced with actual logic
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Hello Aspose.Cells");
                }

                // Save the workbook as HTML to a file
                const string outputPath = "Workbook.html";
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook successfully saved as HTML to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error converting workbook to HTML: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var downloader = new ExcelToHtmlDownloader();
            downloader.DownloadWorkbookAsHtml();
        }
    }
}
