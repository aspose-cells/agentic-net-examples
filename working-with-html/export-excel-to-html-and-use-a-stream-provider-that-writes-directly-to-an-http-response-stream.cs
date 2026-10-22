// Title: Export an Excel workbook to HTML and stream it directly to an HTTP response using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook with Aspose.Cells and saves it as HTML to a provided Stream, such as HttpResponse.Body, with images encoded in Base64 and UTF‑8 encoding. | Show how to set up HtmlSaveOptions for Aspose.Cells to include images inline and write the HTML output directly to an output stream without creating a temporary file.
// Common Searches: Aspose.Cells export Excel to HTML and write to HttpResponse stream in ASP.NET Core | C# save workbook as HTML to a Stream using Aspose.Cells HtmlSaveOptions | How to include images inline when converting Excel to HTML with Aspose.Cells | Stream Aspose.Cells HTML output directly to response without temporary file | Export Excel file to HTML on the fly in a web API using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions inline images | export Excel to HTML stream .NET | write Aspose.Cells HTML to HttpResponse.Body | C# stream provider Excel to HTML conversion | save workbook as HTML directly to stream

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64 and use UTF‑8 encoding, and saves the workbook as HTML directly to a supplied Stream (e.g., HttpResponse.Body), eliminating the need for intermediate files.
public class ExcelExportService
{
    /// <param name="outputStream">The stream to write HTML to (e.g., HttpResponse.Body).</param>
    /// <param name="excelFilePath">Full path to the source Excel file.</param>
    public void ExportExcelToHtml(Stream outputStream, string excelFilePath)
    {
        if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
        if (string.IsNullOrWhiteSpace(excelFilePath)) throw new ArgumentException("Excel file path is required.", nameof(excelFilePath));

        // Prevent FileNotFoundException by checking file existence first.
        if (!File.Exists(excelFilePath))
            throw new FileNotFoundException("The specified Excel file was not found.", excelFilePath);

        try
        {
            // Load the workbook from the specified file path.
            Workbook workbook = new Workbook(excelFilePath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,   // Embed images directly in HTML.
                // ExportSingleSheet property is not available in this version; default behavior exports all sheets.
                Encoding = Encoding.UTF8
            };

            // Write the HTML directly to the provided stream.
            workbook.Save(outputStream, saveOptions);
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context.
            throw new InvalidOperationException("Failed to export Excel to HTML.", ex);
        }
    }
}

// Minimal entry point to satisfy the console application requirement.
public class Program
{
    public static void Main(string[] args)
    {
        // Expect the first argument to be the path of the Excel file.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: <exe> <excelFilePath>");
            return;
        }

        string excelPath = args[0];

        // Verify the Excel file exists before proceeding.
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file '{excelPath}' does not exist.");
            return;
        }

        // Export the Excel file to HTML and write to output.html in the current directory.
        try
        {
            using (FileStream outputStream = new FileStream("output.html", FileMode.Create, FileAccess.Write))
            {
                var service = new ExcelExportService();
                service.ExportExcelToHtml(outputStream, excelPath);
            }

            Console.WriteLine("Export completed successfully. Output saved to 'output.html'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Export failed: {ex.Message}");
        }
    }
}
