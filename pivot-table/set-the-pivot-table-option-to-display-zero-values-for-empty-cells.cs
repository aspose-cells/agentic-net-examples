// Title: How to enable zero-value display for empty cells in an Aspose.Cells pivot table using C# (property not available in current version)
// Common Searches: Aspose.Cells C# show zeros for empty cells in a pivot table | set pivot table to display zero values with Aspose.Cells .NET | Aspose.Cells pivot table option to treat blanks as zero in C# | why is DisplayZeroValues missing in Aspose.Cells pivot tables | example code for configuring empty cell display in Aspose.Cells pivot tables
// Tags: Aspose.Cells pivot table display zero values | C# Aspose.Cells pivot table configuration | Excel pivot table blanks as zero Aspose.Cells | missing DisplayZeroValues property Aspose.Cells | configure empty cell handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, checks the first worksheet for a pivot table, attempts to access the DisplayZeroValues setting (which is not present in the current Aspose.Cells version), logs the limitation, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for at least one pivot table
            if (worksheet.PivotTables.Count > 0)
            {
                // Access the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Note: The DisplayZeroValues property is not available in this version of Aspose.Cells.
                // If needed, configure other pivot table settings here.
                Console.WriteLine("Pivot table found. (Zero-value display setting not supported in this version.)");
            }
            else
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
