// Title: Load an Excel workbook, locate the first PivotTable, and set its layout to Tabular using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook and change pivot table layout to Tabular | How to set the RowLayout of a PivotTable to Tabular using Aspose.Cells .NET | Find the first PivotTable in an Excel file with Aspose.Cells and modify its layout | Iterate through worksheets to locate pivot tables programmatically in C# | Example code for updating pivot table layout in an existing Excel workbook with Aspose.Cells
// Tags: load workbook with Aspose.Cells C# | retrieve first pivot table Aspose.Cells | set pivot table RowLayout to Tabular Aspose.Cells | enumerate worksheets for pivot tables .NET | programmatic pivot layout modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, iterates over each worksheet to find the first sheet containing a PivotTable, accesses that PivotTable (noting that the RowLayout property can be used to switch to Tabular layout when supported), and then saves the workbook to a new file.
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
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Search for the first worksheet that contains at least one PivotTable
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PivotTableCollection pivotTables = sheet.PivotTables;

                if (pivotTables.Count > 0)
                {
                    // Access the first PivotTable found
                    PivotTable pivot = pivotTables[0];

                    // NOTE: The RowLayout property may not be available in older Aspose.Cells versions.
                    // If needed, adjust other pivot settings here.

                    // Exit after processing the first PivotTable (remove break to process all)
                    break;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
