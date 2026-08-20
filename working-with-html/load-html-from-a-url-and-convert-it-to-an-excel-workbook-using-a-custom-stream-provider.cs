// Title: C# – Convert HTML from a URL to Excel with Aspose.Cells using a MemoryStream
// Description: Download an HTML page with HttpClient, load the byte array into a MemoryStream, create an Aspose.Cells Workbook via HtmlLoadOptions, and save it as an XLSX file while ensuring the output folder exists.
// Keywords: Aspose.Cells HTML to Excel C# | load HTML from stream Aspose | HttpClient download HTML C# | HtmlLoadOptions memory stream example | save workbook as XLSX Aspose | convert web page to Excel programmatically | C# Excel generation from HTML
// Common Searches: load html from url into Aspose.Cells workbook | convert online html page to xlsx c# | aspnet download html and save as excel | aspose.cells memory stream html example | c# convert web report to excel file
// Developer Intent: Fetch HTML from a remote URL, stream it into Aspose.Cells, and export the content as an Excel workbook.
// Use Cases: Automated extraction of HTML reports from a web service and conversion to XLSX for downstream analytics. | Scheduled job that transforms dynamically generated web pages into Excel spreadsheets for business users. | Integration of HTML‑to‑Excel conversion into a CI/CD pipeline that stores results in a predefined directory.
// AI Prompts: Generate C# code that downloads HTML via HttpClient, loads it into a MemoryStream, and saves it as an XLSX file using Aspose.Cells. | Explain how to configure HtmlLoadOptions to retain CSS styling when converting a web page to Excel with Aspose.Cells. | Show how to refactor the sample to use async/await for HttpClient calls and proper exception handling.

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// Download an HTML page with HttpClient, load the byte array into a MemoryStream, create an Aspose.Cells Workbook via HtmlLoadOptions, and save it as an XLSX file while ensuring the output folder exists.
public class HtmlToExcelConverter
{
    public static void Run()
    {
        try
        {
            // URL of the HTML file to be converted.
            const string htmlUrl = "https://example.com/sample.html";

            // Download the HTML content.
            byte[] htmlData;
            using (HttpClient client = new HttpClient())
            {
                htmlData = client.GetByteArrayAsync(htmlUrl).Result;
            }

            // Load the HTML content into a workbook from a memory stream.
            using (MemoryStream htmlStream = new MemoryStream(htmlData))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Define the output Excel file path.
                const string outputPath = "ConvertedFromHtml.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an Excel file.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Conversion successful. File saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        HtmlToExcelConverter.Run();
    }
}
