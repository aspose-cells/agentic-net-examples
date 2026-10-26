// Title: How to disable auto‑fit of column widths for a PivotTable using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# prevent pivot table columns from auto‑fitting | how to set a fixed column width for a pivot table using Aspose.Cells | disable automatic column width adjustment in an Excel pivot table with .NET | C# code example to keep custom column sizes in a pivot table after saving | Aspose.Cells pivot table column width control without AutoFitColumnWidth property
// Tags: Aspose.Cells pivot table column width control | C# set fixed width for Excel pivot table columns | disable auto‑fit for pivot table columns in Aspose.Cells | prevent automatic column resizing in Excel pivot tables | save workbook after adjusting pivot table column size

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, locates the first PivotTable, and sets a fixed width for its first column using Worksheet.Cells.SetColumnWidth, effectively disabling the auto‑fit behavior before saving the modified file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet where the PivotTable resides (adjust the name as needed)
            Worksheet worksheet = workbook.Worksheets["Sheet1"];
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet \"Sheet1\" not found.");
                return;
            }

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // (Optional) Disable automatic column width fitting if supported
            // pivotTable.AutoFitColumnWidth = false; // Not available in some versions

            // Set a custom width for the first column of the PivotTable
            int firstColumnIndex = pivotTable.ColumnRange.StartColumn; // Use StartColumn for compatibility
            worksheet.Cells.SetColumnWidth(firstColumnIndex, 15); // width in characters

            // Save the workbook with the changes
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
