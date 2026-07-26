// Title: C# – Convert HTML from a URL to Excel using Aspose.Cells and MemoryStream
// Description: Download an HTML page with HttpClient, feed the byte array into a MemoryStream, load it into an Aspose.Cells Workbook via HtmlLoadOptions, and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# HTML to Excel | MemoryStream | HtmlLoadOptions | download HTML | convert web page to XLSX | load HTML from URL | Workbook.Save | HttpClient | stream provider
// Common Searches: Aspose.Cells load HTML from URL C# | Convert web page to Excel using MemoryStream | C# download HTML and save as XLSX | How to use HtmlLoadOptions with a stream in Aspose.Cells | Example converting remote HTML to Excel in .NET
// Developer Intent: Import HTML retrieved from a remote URL into an Aspose.Cells Workbook via a stream and export it as an Excel file.
// Use Cases: Automate daily reporting by pulling an HTML dashboard from a web service and converting it to XLSX. | Extract tables from external web pages and store them in Excel for downstream analysis. | Expose an API endpoint that accepts a URL, transforms the remote HTML into an Excel workbook, and returns the file.
// AI Prompts: Generate C# code that downloads HTML from a given URL, loads it into an Aspose.Cells Workbook using HtmlLoadOptions and a MemoryStream, then saves the workbook as an XLSX file. | Explain strategies for handling large HTML files during conversion with Aspose.Cells, focusing on stream buffering and memory usage. | Provide best‑practice error‑handling patterns for converting remote HTML pages to Excel using Aspose.Cells in .NET.

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// Download an HTML page with HttpClient, feed the byte array into a MemoryStream, load it into an Aspose.Cells Workbook via HtmlLoadOptions, and save the result as an XLSX file.
public class Program
{
    public static void Main()
    {
        try
        {
            // URL of the HTML page to be converted.
            string htmlUrl = "https://example.com/sample.html";

            // Download the HTML content.
            byte[] htmlData;
            using (HttpClient client = new HttpClient())
            {
                htmlData = client.GetByteArrayAsync(htmlUrl).Result;
            }

            // Load the HTML into a workbook from a memory stream.
            using (MemoryStream stream = new MemoryStream(htmlData))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(stream, loadOptions);

                // Save the workbook as an Excel file.
                string outputPath = "Converted.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
