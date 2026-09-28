// Title: How to stream an Aspose.Cells workbook as an inline XLSX file via HttpResponse in ASP.NET (C#)
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, adds sample data, sets the HTTP response header Content‑Disposition to inline, and streams the workbook in XLSX format directly to the HttpResponse output stream. | Show an ASP.NET MVC controller action that uses Aspose.Cells to export a worksheet as an inline Excel file, including proper error handling and response cleanup. | Provide a minimal ASP.NET Core endpoint that returns an Aspose.Cells‑generated XLSX file with Content‑Disposition set to inline, using Workbook.Save with SaveFormat.Xlsx.
// Common Searches: aspocells inline excel download asp.net mvc | c# set content-disposition inline for aspose.cells generated xlsx | stream workbook to HttpResponse outputstream using Aspose.Cells | asp.net core return aspose.cells workbook as inline file | how to export Aspose.Cells workbook without attachment header
// Tags: Aspose.Cells save workbook to HttpResponse inline | Aspose.Cells stream XLSX to ASP.NET response | Aspose.Cells set Content-Disposition inline | ASP.NET MVC export Excel with Aspose.Cells | ASP.NET Core inline Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a new Aspose.Cells Workbook, writing a value to cell A1, and saving the workbook in XLSX format to a supplied Stream. By passing the HttpResponse.OutputStream to the ExportWorkbook method and setting the response header Content‑Disposition to "inline", the workbook can be streamed directly to the client as an inline Excel file. The code includes argument validation and exception wrapping for robust error handling.
public class WorkbookExportHandler
{
    /// <param name="outputStream">The stream to which the workbook will be written.</param>
    public void ExportWorkbook(Stream outputStream)
    {
        if (outputStream == null)
            throw new ArgumentNullException(nameof(outputStream));

        try
        {
            // Create a new workbook (default format is XLSX)
            var workbook = new Workbook();

            // Add sample data to the first worksheet
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Save the workbook directly to the provided stream in XLSX format
            workbook.Save(outputStream, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to allow callers to handle the failure
            throw new InvalidOperationException("Failed to export workbook.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Define output file path
        string outputPath = "ExportedWorkbook.xlsx";

        try
        {
            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Export workbook to file
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                var handler = new WorkbookExportHandler();
                handler.ExportWorkbook(fileStream);
            }

            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error exporting workbook: {ex.Message}");
        }
    }
}
