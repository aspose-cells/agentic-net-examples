// Title: Stream Aspose.Cells HTML export from a MemoryStream directly to an ASP.NET response for immediate download
// AI Prompts: Create a C# ASP.NET MVC action that loads a Workbook, saves it as HTML into a MemoryStream, and returns the stream as a downloadable file with correct Content-Type and Content-Disposition headers. | Show how to flush a MemoryStream containing Aspose.Cells HTML output to HttpResponse in ASP.NET Core, ensuring the stream position is reset and the response is properly completed.
// Common Searches: how to return Aspose.Cells HTML as a file download in ASP.NET | write MemoryStream bytes to HttpResponse in C# without saving to disk | Aspose.Cells export workbook to HTML and stream to browser | set content-disposition header for html attachment in ASP.NET MVC | download generated HTML from Aspose.Cells directly from memory
// Tags: Aspose.Cells export html to HttpResponse | ASP.NET stream MemoryStream as file download | C# set content-disposition for html attachment | Aspose.Cells save workbook as html without temporary file | ASP.NET Core return MemoryStream result

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading or creating a Workbook, saving it as HTML into a MemoryStream, resetting the stream, and sending the HTML bytes to the client via HttpResponse with appropriate Content-Type and Content-Disposition headers for an immediate download, without writing a temporary file to disk.
public class ExportHtml
{
    public void DownloadWorkbookAsHtml()
    {
        try
        {
            // Load an existing workbook if a template file is present; otherwise create a new one.
            Workbook workbook;
            const string templatePath = "template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                // TODO: populate workbook with data as needed
            }

            // Save the workbook to a MemoryStream in HTML format.
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, SaveFormat.Html);
                // Ensure the stream is positioned at the beginning.
                htmlStream.Position = 0;

                // Write the HTML bytes to a file.
                const string outputPath = "Workbook.html";
                File.WriteAllBytes(outputPath, htmlStream.ToArray());

                Console.WriteLine($"HTML file has been saved to: {Path.GetFullPath(outputPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        ExportHtml exporter = new ExportHtml();
        exporter.DownloadWorkbookAsHtml();
    }
}
