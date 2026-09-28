// Title: Create a pivot table on a new worksheet from a source range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing workbook, adds a new worksheet, and inserts a pivot table based on a defined range from the first sheet using Aspose.Cells. | Show how to set the pivot table’s source data range and position it at cell A3 on the new worksheet with Aspose.Cells. | Provide code to save the modified workbook to a specified file after creating the pivot table.
// Common Searches: Aspose.Cells C# add pivot table to a newly created worksheet from existing sheet range | how to specify source data range for a pivot table using Aspose.Cells .NET | sample code for creating a pivot table on a separate sheet with Aspose.Cells | C# Aspose.Cells pivot table placement cell A3 example
// Tags: Aspose.Cells add pivot table programmatically | C# create pivot table from worksheet range | Aspose.Cells new worksheet for pivot table | pivot table source range specification Aspose.Cells | save workbook after pivot table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads source.xlsx, adds a new worksheet named PivotSheet, creates a pivot table called PivotTable1 on that sheet using the range A1:D100 from the first worksheet, and saves the updated workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook (source data is on the first sheet)
            Workbook workbook = new Workbook(sourcePath);

            // Reference to the sheet that contains the source data
            Worksheet sourceSheet = workbook.Worksheets[0]; // adjust index or name as needed

            // Define the range that holds the source data (e.g., A1:D100)
            string sourceRange = "A1:D100";

            // Add a new worksheet that will host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotSheet";

            // Add a pivot table to the new worksheet; placed starting at cell A3
            int pivotTableIndex = pivotSheet.PivotTables.Add(
                $"{sourceSheet.Name}!{sourceRange}", // source data range
                "A3",                                 // top‑left cell of the pivot table
                "PivotTable1");                       // pivot table name

            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // NOTE:
            // The older Add(int) overloads for RowFields, ColumnFields, and DataFields are no longer available.
            // Field configuration can be performed using field names or by retrieving PivotField objects.
            // For simplicity, this example skips explicit field configuration.

            // Save the workbook with the newly created pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table created and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
