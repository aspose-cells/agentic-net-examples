// Title: How to delete a specific pivot table by name in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove pivot table by its name from workbook | Delete specific pivot table using RemoveAt in Aspose.Cells | C# code to find and delete a pivot table in an Excel file with Aspose | Programmatically delete a pivot table in .NET using Aspose.Cells | Remove pivot table from worksheet collection Aspose.Cells example
// Tags: Aspose.Cells delete pivot table | PivotTableCollection RemoveAt C# | remove named pivot table Aspose | Excel workbook pivot table deletion .NET | C# delete pivot table by name

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The DeletePivotTable utility validates the source file, loads it with Aspose.Cells, searches each worksheet's PivotTableCollection for a pivot table matching the supplied name (case‑insensitive), removes it using the collection's RemoveAt method, and saves the modified workbook to the target path, throwing clear exceptions when the file is missing or the pivot table cannot be found.
public static class PivotTableUtility
{
    /// <param name="inputFilePath">Full path to the source Excel file.</param>
    /// <param name="pivotTableName">Name of the pivot table to delete.</param>
    /// <param name="outputFilePath">Full path where the modified workbook will be saved.</param>
    public static void DeletePivotTable(string inputFilePath, string pivotTableName, string outputFilePath)
    {
        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFilePath))
            throw new FileNotFoundException($"Input file not found: {inputFilePath}");

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFilePath);
            bool pivotFound = false;

            // Iterate through all worksheets to locate the pivot table
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Search for the pivot table by name
                for (int i = 0; i < pivots.Count; i++)
                {
                    if (string.Equals(pivots[i].Name, pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the pivot table using its index
                        pivots.RemoveAt(i);
                        pivotFound = true;
                        break; // Exit inner loop after removal
                    }
                }

                if (pivotFound)
                    break; // Exit outer loop once the pivot table is deleted
            }

            if (!pivotFound)
                throw new ArgumentException($"Pivot table '{pivotTableName}' was not found in the workbook.");

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace while providing context
            throw new InvalidOperationException("Failed to delete pivot table.", ex);
        }
    }

    // Entry point for testing the utility
    public static void Main(string[] args)
    {
        try
        {
            // Example usage – replace with actual paths and pivot table name
            string inputPath = @"C:\Temp\InputWorkbook.xlsx";
            string pivotName = "PivotTable1";
            string outputPath = @"C:\Temp\OutputWorkbook.xlsx";

            DeletePivotTable(inputPath, pivotName, outputPath);
            Console.WriteLine("Pivot table deleted and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
