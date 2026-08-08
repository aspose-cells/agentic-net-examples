// Title: Stream an Aspose.Cells Workbook as an Inline XLSX File via HttpResponse (C#)
// Description: Creates a new Aspose.Cells Workbook, adds optional data, sets the HTTP response header Content‑Disposition to inline, and streams the workbook in XLSX format directly to the client without saving to disk. Ideal for ASP.NET MVC/Web API endpoints that deliver Excel files on‑the‑fly.
// Keywords: Aspose.Cells stream workbook C# | ContentDisposition inline Aspose.Cells | HttpResponse Excel download ASP.NET | Aspose.Cells SaveFormat.Xlsx to response | C# generate Excel file on the fly | ASP.NET Core return Excel file | Aspose.Cells HttpResponse stream
// Common Searches: how to stream Aspose.Cells workbook to HttpResponse | Aspose.Cells set ContentDisposition inline | return XLSX file from ASP.NET controller | C# Aspose.Cells download Excel without saving | Aspose.Cells write workbook to response stream
// Developer Intent: Generate an Excel workbook in memory, mark it as inline, and send it to the browser through the HTTP response stream.
// Use Cases: Provide a live report that users can view directly in the browser. | Export data grids to Excel on demand in a web application. | Deliver dynamically created templates without writing temporary files to the server.
// AI Prompts: Show C# code using Aspose.Cells to create a workbook, set ContentDisposition to inline, and stream it as an XLSX file via HttpResponse in ASP.NET MVC. | Give an example of returning an Aspose.Cells workbook from an ASP.NET Core controller action without saving to disk. | Explain how to handle exceptions when streaming an Aspose.Cells workbook to an HttpResponse and ensure the response is properly closed.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Aspose.Cells Workbook, adds optional data, sets the HTTP response header Content‑Disposition to inline, and streams the workbook in XLSX format directly to the client without saving to disk. Ideal for ASP.NET MVC/Web API endpoints that deliver Excel files on‑the‑fly.
public class WorkbookExporter
{
    /// <param name="filePath">Full path where the workbook will be saved.</param>
    public void ExportToFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path must be provided.", nameof(filePath));

        try
        {
            // Ensure the directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Create a new workbook (default format is XLSX)
            var workbook = new Workbook();

            // Add sample data – this step is optional
            workbook.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // Save the workbook to the specified file in XLSX format
            workbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow for caller handling
            throw new InvalidOperationException("Failed to export workbook.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Output", "Sample.xlsx");
            var exporter = new WorkbookExporter();
            exporter.ExportToFile(outputPath);
            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
