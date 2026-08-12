// Title: Stream Aspose.Cells HTML Export Bytes Directly to ASP.NET Response for Instant Download
// Description: Export an Aspose.Cells workbook to HTML in a MemoryStream and write the resulting byte array to an ASP.NET HttpResponse (MVC, Web Forms, or Web API). The guide shows how to set the correct MIME type and Content‑Disposition headers so the browser downloads the HTML file without creating a temporary file on the server.
// Keywords: Aspose.Cells | HTML export | MemoryStream | ASP.NET response | file download | C# streaming | HttpResponse | FileResult | Web API | Web Forms
// Common Searches: Aspose.Cells export HTML to HttpResponse | download generated HTML in ASP.NET MVC | write byte[] to ASP.NET response stream | ASP.NET file download from MemoryStream | C# stream Aspose.Cells HTML without temp file
// Developer Intent: Send the HTML byte array produced by Aspose.Cells to the client as a downloadable file in a single HTTP response.
// Use Cases: One‑click "Export to HTML" button in an ASP.NET MVC controller that streams the file to the browser. | Web Forms page that generates a workbook on the fly and returns the HTML as an attachment. | Web API endpoint that returns the exported HTML using FileResult or IActionResult without writing to disk.
// AI Prompts: Create an ASP.NET MVC action that calls ExportWorkbookAsHtml, sets Content‑Type to text/html, adds a Content‑Disposition attachment header, and returns a FileResult. | Show a Web Forms Page_Load example that writes the HTML byte array to Response.OutputStream and ends the response. | Provide a minimal ASP.NET Core Web API method that streams Aspose.Cells HTML bytes as a downloadable file using IActionResult.

using System;
using System.IO;
using Aspose.Cells;

// Export an Aspose.Cells workbook to HTML in a MemoryStream and write the resulting byte array to an ASP.NET HttpResponse (MVC, Web Forms, or Web API). The guide shows how to set the correct MIME type and Content‑Disposition headers so the browser downloads the HTML file without creating a temporary file on the server.
public class HtmlExportHelper
{
    // Exports a workbook as an HTML file and returns the HTML bytes.
    public byte[] ExportWorkbookAsHtml()
    {
        Workbook workbook = null;
        try
        {
            // Create a new workbook and add some sample data.
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Save the workbook to a memory stream in HTML format.
            using (var htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, SaveFormat.Html);
                return htmlStream.ToArray(); // Return the generated HTML bytes.
            }
        }
        catch (Exception ex)
        {
            // Wrap and rethrow the exception for caller handling.
            throw new InvalidOperationException("Failed to export workbook as HTML.", ex);
        }
        finally
        {
            // Ensure the workbook is properly disposed.
            workbook?.Dispose();
        }
    }
}

public class Program
{
    // Entry point required for compilation.
    public static void Main()
    {
        try
        {
            var helper = new HtmlExportHelper();
            byte[] htmlBytes = helper.ExportWorkbookAsHtml();

            // Write the HTML bytes to a file for verification.
            string outputPath = "output.html";
            File.WriteAllBytes(outputPath, htmlBytes);
            Console.WriteLine($"HTML exported successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
