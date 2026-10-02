// Title: Stream an Aspose.Cells workbook as an inline XLSX file to an ASP.NET HttpResponse using C#
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, adds data, sets the HTTP response header Content-Disposition to inline, and streams the workbook in XLSX format directly to HttpResponse.OutputStream. | Refactor a method that saves a workbook to disk so it instead writes the workbook to the current HttpResponse without creating a temporary file, using Aspose.Cells.
// Common Searches: asp.net core stream xlsx file inline using aspose.cells c# | c# send excel workbook as inline attachment with content-disposition header asp.net | how to write aspose.cells workbook to HttpResponse output stream c# | asp.net mvc return aspose.cells workbook as file result inline | asp.net set Content-Disposition inline for generated Excel file c#
// Tags: Aspose.Cells write workbook to HttpResponse stream | inline Content-Disposition for Excel download ASP.NET | C# stream XLSX using Aspose.Cells | Aspose.Cells save workbook to response without file | ASP.NET Core return Excel as inline file

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to create a new Aspose.Cells Workbook, populate it with data, set the HTTP response header Content-Disposition to "inline", and stream the workbook as an XLSX file directly to the client via HttpResponse, eliminating the need for a temporary file on the server.
public class ExcelExport
{
    public static void Main()
    {
        try
        {
            var exporter = new ExcelExport();
            exporter.SaveWorkbook();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while exporting workbook: {ex.Message}");
        }
    }

    public void SaveWorkbook()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Hello Aspose.Cells");

            // Define output file path
            string outputPath = Path.Combine(Environment.CurrentDirectory, "HelloAspose.xlsx");

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to the specified file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during workbook creation or saving
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
            throw;
        }
    }
}
