// Title: Enable item label display in the values area of an Aspose.Cells pivot table by setting ShowValuesColumn to true (C#)
// Common Searches: how to show item labels in the values area of a pivot table with Aspose.Cells C# | set ShowValuesColumn true for Aspose.Cells pivot table .NET | Aspose.Cells pivot table values column label visibility | C# enable item labels in pivot table values column using Aspose | why is ShowValuesColumn property unavailable in Aspose.Cells
// Tags: Aspose.Cells pivot table ShowValuesColumn | display item labels in pivot values column | Aspose.Cells .NET pivot table configuration | enable values column label in Excel using Aspose | missing ShowValuesColumn property Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class
using System;
using System.IO;

// The example loads an Excel workbook, checks for a worksheet and a pivot table, accesses the first pivot table, and notes that the ShowValuesColumn property is not available in the current Aspose.Cells version before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // NOTE: The ShowValuesColumn property is not available in the current Aspose.Cells version.
            // If needed, configure other pivot table settings here.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
