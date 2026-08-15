// Title: Aspose.Cells .NET – Set StandardWidth and AutoFit a Column Range (Override Example)
// Description: Shows how to define a worksheet's default column width, manually adjust a single column, auto‑fit a selected range, read column widths before and after the operation, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells StandardWidth | AutoFitColumns range | SetColumnWidth before AutoFit | GetColumnWidth after AutoFit | .NET spreadsheet column width | default column width Aspose | override column width | column width debugging Aspose.Cells
// Common Searches: Aspose.Cells set default column width | AutoFitColumns ignore manual width | How to get column width after AutoFit | C# Aspose.Cells column width range | Preserve custom column width while auto‑fitting
// Developer Intent: Define a workbook's standard column width, apply a custom width to one column, auto‑fit a specific range, and confirm the resulting widths programmatically.
// Use Cases: Create a template where most columns follow a standard width but selected columns expand to fit their content. | Generate a report that keeps a particular column narrow for layout constraints while allowing other columns to auto‑adjust. | Debug column‑width behavior by logging widths before and after AutoFit to ensure manual overrides are respected.
// AI Prompts: Write C# code that sets cells.StandardWidth, overrides column B with SetColumnWidth, auto‑fits columns A‑D, and prints widths before and after using Aspose.Cells. | Explain the interaction between AutoFitColumns and manually set column widths in Aspose.Cells for .NET, and how to keep a custom width while auto‑fitting other columns. | Show how to retrieve column widths programmatically, compare them before and after applying AutoFitColumns to a defined range, and save the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to define a worksheet's default column width, manually adjust a single column, auto‑fit a selected range, read column widths before and after the operation, and save the workbook with Aspose.Cells for .NET.
    public class StandardWidthAutoFitDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set the default column width (standard width) in characters
            cells.StandardWidth = 20.0; // default width for all columns

            // Populate sample data in columns A to D
            cells["A1"].PutValue("Short");
            cells["B1"].PutValue("This is a longer text that will need more width");
            cells["C1"].PutValue("Medium length");
            cells["D1"].PutValue("Very very long text that definitely exceeds the standard width");

            // Override width of column B (index 1) before autofit
            cells.SetColumnWidth(1, 10.0); // custom width for column B

            // Show column widths before autofit
            Console.WriteLine("Column widths before AutoFit:");
            for (int i = 0; i < 4; i++)
            {
                Console.WriteLine($"Column {i} width: {cells.GetColumnWidth(i)}");
            }

            // AutoFit columns A through D (indices 0 to 3)
            worksheet.AutoFitColumns(0, 3);

            // Show column widths after autofit
            Console.WriteLine("Column widths after AutoFit:");
            for (int i = 0; i < 4; i++)
            {
                Console.WriteLine($"Column {i} width: {cells.GetColumnWidth(i)}");
            }

            // Save the workbook
            string outputPath = "StandardWidthAutoFitDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
