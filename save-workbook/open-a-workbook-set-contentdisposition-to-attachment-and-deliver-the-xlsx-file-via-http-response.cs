// Title: How to stream an existing Excel workbook as an XLSX attachment in an ASP.NET Core response using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook from a file path with Aspose.Cells and returns it as a downloadable XLSX file from an ASP.NET Core controller, setting the Content‑Disposition header to attachment. | Show how to adapt the ExportAsAttachment method to accept an HttpResponse object and write the workbook directly to HttpResponse.Body in an ASP.NET Core MVC action. | Provide a minimal ASP.NET Core endpoint that uses Aspose.Cells to open a .xlsx file and streams it to the client as an attachment with the correct MIME type.
// Common Searches: asp.net core download excel file as attachment using aspose.cells SaveFormat.Xlsx | c# Aspose.Cells stream workbook to HttpResponse.Body with content-disposition header | how to return generated Excel workbook from ASP.NET Core API as file download | Aspose.Cells export existing workbook to HTTP response without saving to disk | set Content-Disposition: attachment for Excel file in ASP.NET Core controller using Aspose.Cells
// Tags: Aspose.Cells export workbook to HTTP response | stream XLSX file as attachment in ASP.NET Core | Content-Disposition header for Excel download | SaveFormat.Xlsx with Aspose.Cells | C# load workbook from file path Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example defines a WorkbookExporter class that validates inputs, loads an existing workbook with Aspose.Cells, and saves it in XLSX format to a supplied stream. By passing HttpResponse.Body as the output stream, the workbook can be streamed directly to the client as a downloadable attachment, with the appropriate Content‑Disposition header set in an ASP.NET Core controller.
public class WorkbookExporter
{
    /// <param name="outputStream">The stream to which the workbook will be written (e.g., HttpResponse.Body).</param>
    /// <param name="filePath">Full path to the workbook to be opened.</param>
    public void ExportAsAttachment(Stream outputStream, string filePath)
    {
        if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path must be provided.", nameof(filePath));

        try
        {
            // Verify that the workbook file exists.
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Workbook file not found.", filePath);

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Write the workbook to the output stream in XLSX format.
            workbook.Save(outputStream, SaveFormat.Xlsx);
        }
        catch (Exception)
        {
            // Rethrow after any necessary logging or handling.
            throw;
        }
    }
}

public class Program
{
    // Usage: dotnet run <inputWorkbookPath> <outputFilePath>
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <inputWorkbookPath> <outputFilePath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            // Ensure the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found: {inputPath}");
                return;
            }

            // Create or overwrite the output file.
            using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                WorkbookExporter exporter = new WorkbookExporter();
                exporter.ExportAsAttachment(outputStream, inputPath);
            }

            Console.WriteLine($"Workbook exported successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
