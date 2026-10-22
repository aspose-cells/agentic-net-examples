// Title: Delete the first PivotTable from an Excel worksheet using Aspose.Cells PivotTables.RemoveAt in C#
// Common Searches: C# code to remove a pivot table from an Excel worksheet with Aspose.Cells | Aspose.Cells PivotTables.RemoveAt example for deleting the first pivot table | How to check for and delete pivot tables in a workbook using Aspose.Cells .NET | Remove a specific pivot table by index in an Excel file with Aspose.Cells | Delete all pivot tables from a worksheet using Aspose.Cells C# API
// Tags: Aspose.Cells PivotTables.RemoveAt method usage | C# delete pivot table from Excel worksheet | remove pivot table by index Aspose.Cells | Excel workbook pivot table removal with .NET | handling pivot tables in Aspose.Cells applications | Aspose.Cells worksheet pivot table management

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file with Aspose.Cells, checks if the first worksheet contains any PivotTables, removes the first PivotTable using worksheet.PivotTables.RemoveAt(0), and saves the modified workbook, while handling CellsException and generic errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied). If the file is password‑protected,
            // a CellsException will be thrown and handled in the catch block.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove the first PivotTable if any exist
            if (worksheet.PivotTables.Count > 0)
            {
                worksheet.PivotTables.RemoveAt(0);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors specific to Aspose.Cells (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
