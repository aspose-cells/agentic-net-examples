// Title: C# utility to extract each worksheet’s paper width and height (in inches) from an Excel file using Aspose.Cells and return a JSON summary
// Common Searches: how to get worksheet page setup paper dimensions in inches with Aspose.Cells for .NET | C# read paper width and height of all sheets in an Excel workbook using Aspose.Cells | generate JSON report of Excel sheet paper sizes using Aspose.Cells library | convert Aspose.Cells PaperWidth points to inches in C# code | list worksheet names with their page setup paper size via Aspose.Cells
// Tags: Aspose.Cells worksheet page setup extraction | paper size conversion points to inches C# | JSON serialization of worksheet dimensions | iterate workbook worksheets Aspose.Cells | extract Excel sheet paper dimensions .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The utility loads an Excel workbook with Aspose.Cells, iterates through every worksheet, reads the PageSetup.PaperWidth and PaperHeight values (provided in points), converts them to inches, records each sheet's name and dimensions in a list of WorksheetPaperSizeInfo objects, and returns the collection as a formatted JSON string.
public class WorksheetPaperSizeInfo
{
    // Initialize to empty to satisfy non‑nullable warning.
    public string WorksheetName { get; set; } = string.Empty;
    public double PaperWidthInInches { get; set; }
    public double PaperHeightInInches { get; set; }
}

public static class PaperSizeUtility
{
    // Extracts paper width and height (in inches) from all worksheets and returns a JSON summary.
    public static string GetPaperSizesJson(string workbookPath)
    {
        // Verify the file exists before attempting to load.
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        // Load the workbook from the specified file.
        var workbook = new Workbook(workbookPath);

        var summary = new List<WorksheetPaperSizeInfo>();

        // Iterate through each worksheet in the workbook.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var pageSetup = sheet.PageSetup;

            // PaperWidth and PaperHeight are returned in points (1 point = 1/72 inch).
            double widthInInches = pageSetup.PaperWidth / 72.0;
            double heightInInches = pageSetup.PaperHeight / 72.0;

            summary.Add(new WorksheetPaperSizeInfo
            {
                WorksheetName = sheet.Name,
                PaperWidthInInches = widthInInches,
                PaperHeightInInches = heightInInches
            });
        }

        // Serialize the summary list to a formatted JSON string.
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(summary, jsonOptions);
    }
}

public class Program
{
    // Entry point required for console application.
    public static void Main(string[] args)
    {
        try
        {
            // Expect the workbook path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: dotnet run <workbookPath>");
                return;
            }

            string workbookPath = args[0];

            // Get JSON summary of paper sizes.
            string jsonResult = PaperSizeUtility.GetPaperSizesJson(workbookPath);

            // Output the result.
            Console.WriteLine(jsonResult);
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
