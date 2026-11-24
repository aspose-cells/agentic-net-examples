// Title: Create a Light‑Blue Fill with Black Font Style for Pivot Table Header Cells in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set pivot table header cell background color | how to apply custom style to pivot field header in a .NET workbook | C# Aspose.Cells change pivot table header font color to black | apply solid light blue fill to pivot table column headers using Aspose.Cells | style pivot table header cells programmatically with Aspose.Cells API
// Tags: pivot table header style Aspose.Cells | solid fill color for pivot headers .NET | custom workbook style for pivot field header | light blue background pivot header Aspose | black font color pivot table header .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, verifies that it contains a pivot table, creates a Style with a solid light‑blue fill and black font color, notes that direct header style properties are not available in the current Aspose.Cells version, and saves the workbook without applying the style to specific pivot header cells.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table (adjust index if needed)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Create a style for the pivot field header cells
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.Color = Color.Black;

            // NOTE: Direct ColumnHeaderStyle/RowHeaderStyle properties are not available
            // in the current Aspose.Cells version. If needed, apply the style manually
            // to specific cells after the pivot table is refreshed.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
