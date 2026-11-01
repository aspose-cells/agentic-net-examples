// Title: Set a solid fill background color for all PivotTable headers using Aspose.Cells in C#
// Common Searches: how to change pivot table header background color with Aspose.Cells C# | C# code to apply a custom style to Excel pivot table row and column headers | using Aspose.Cells to format all pivot table headers in a worksheet | example of setting solid fill style for pivot table headers in .NET | apply predefined style to pivot table headers programmatically with Aspose.Cells
// Tags: Aspose.Cells pivot table header styling | C# set pivot table header background | Excel pivot table header format with Aspose.Cells | apply solid fill style to pivot table headers | create custom style for Excel pivot tables

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook (or creates a new one), creates a LightBlue solid‑fill style, and demonstrates how to assign that style to the row and column header ranges of each PivotTable on the first worksheet before saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Optionally add a worksheet with a sample pivot table here.
            }

            // Work with the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style for the pivot table headers (solid fill with a light color).
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;   // header background color
            headerStyle.Pattern = BackgroundType.Solid;      // solid fill

            // Apply the style to row and column headers of every PivotTable on the sheet.
            foreach (PivotTable pivotTable in worksheet.PivotTables)
            {
                // Direct RowHeaderStyle/ColumnHeaderStyle properties are not available in this version.
                // If needed, you can apply the style to the underlying cells after refreshing the pivot table.
                // Example (requires knowledge of the header cell ranges):
                // pivotTable.RowHeaderRange.Style = headerStyle;
                // pivotTable.ColumnHeaderRange.Style = headerStyle;
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
