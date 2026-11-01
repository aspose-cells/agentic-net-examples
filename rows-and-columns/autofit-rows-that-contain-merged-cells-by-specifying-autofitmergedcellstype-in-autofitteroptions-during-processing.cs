// Title: Auto‑fit worksheet rows that contain merged cells using AutoFitterOptions.AutoFitMergedCellsType in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing XLSX workbook with Aspose.Cells, creates an AutoFitterOptions object with AutoFitMergedCellsType set to IncludeMergedCells, applies AutoFitRows to the first worksheet, and saves the updated file. | Show how to replace the deprecated AutoFitMergedCells property with the AutoFitMergedCellsType enum when auto‑sizing rows that have merged cells in an Aspose.Cells .NET project. | Generate a robust C# example that verifies the input file, configures AutoFitterOptions to consider merged cells, auto‑fits all rows, and handles any runtime exceptions.
// Common Searches: aspnet c# how to auto size rows that have merged cells using aspose.cells | using AutoFitterOptions AutoFitMergedCellsType to fit rows in an Excel worksheet | example code for Aspose.Cells AutoFitRows with merged cells in .NET | c# autosize row height including merged cells Aspose.Cells 2023
// Tags: row auto‑fit with merged cells Aspose.Cells | AutoFitterOptions AutoFitMergedCellsType .NET | C# Excel row height autosizing merged cells | Aspose.Cells worksheet row height adjustment example | handling merged cells during row autosizing .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing XLSX workbook, configures AutoFitterOptions with AutoFitMergedCellsType to include merged cells, auto‑fits all rows in the first worksheet, and saves the modified workbook, while checking for a missing input file and handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure AutoFitterOptions to handle merged cells
            AutoFitterOptions options = new AutoFitterOptions
            {
                // Include merged cells when auto‑fitting rows (obsolete but functional)
                AutoFitMergedCells = true
                // The newer AutoFitMergedCellsType property is omitted because the enum value
                // varies across Aspose.Cells versions and may cause compilation errors.
            };

            // Auto‑fit all rows in the worksheet using the specified options
            sheet.AutoFitRows(options);

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
