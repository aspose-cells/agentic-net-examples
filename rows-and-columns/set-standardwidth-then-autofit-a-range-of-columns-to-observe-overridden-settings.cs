// Title: Aspose.Cells .NET – Set StandardWidth and AutoFit Columns While Preserving Manual Width Overrides
// Description: Demonstrates how to define a worksheet's default column width (StandardWidth), manually adjust a specific column, and then auto‑fit a range of columns. The example logs column widths before and after AutoFitColumns and saves the workbook, showing that explicit width settings are respected.
// Keywords: Aspose.Cells StandardWidth | AutoFitColumns .NET | override column width Aspose.Cells | default column width worksheet | column width logging Aspose.Cells | C# spreadsheet column sizing
// Common Searches: set default column width then auto‑fit in Aspose.Cells | preserve manually set column width after AutoFitColumns | Aspose.Cells StandardWidth example | how to log column widths before and after autofit Aspose.Cells | C# Aspose.Cells column width override
// Developer Intent: Define a global default column width, change one column manually, and auto‑fit other columns to see how the manual setting is maintained.
// Use Cases: Create templates with a uniform default width but keep certain columns narrow for layout control. | Generate reports where most columns expand automatically while a key column stays fixed. | Validate column‑width behavior by comparing measurements before and after auto‑fit.
// AI Prompts: Write C# code using Aspose.Cells to set StandardWidth to 25, set column B width to 8, auto‑fit columns A‑D, and print the final widths. | Explain the interaction between AutoFitColumns and manually sized columns in Aspose.Cells, including how to retrieve widths after fitting. | Provide a step‑by‑step guide to log column widths before and after calling AutoFitColumns for a selected range in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to define a worksheet's default column width (StandardWidth), manually adjust a specific column, and then auto‑fit a range of columns. The example logs column widths before and after AutoFitColumns and saves the workbook, showing that explicit width settings are respected.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set the default column width (StandardWidth) for the worksheet
        cells.StandardWidth = 20.0; // 20 characters
        Console.WriteLine($"StandardWidth set to: {cells.StandardWidth}");

        // Populate sample data in columns A to D
        cells["A1"].PutValue("Short");
        cells["B1"].PutValue("This is a longer text that should cause column B to expand when auto‑fit is applied");
        cells["C1"].PutValue("Medium length");
        cells["D1"].PutValue("Very very very long text that will definitely exceed the standard width");

        // Manually override the width of column C (index 2) to be narrower than the standard width
        cells.SetColumnWidth(2, 5.0);
        Console.WriteLine($"Column C width before autofit: {cells.GetColumnWidth(2)}");

        // Display column widths before auto‑fit for columns A‑D
        for (int i = 0; i <= 3; i++)
        {
            Console.WriteLine($"Column {(char)('A' + i)} width before autofit: {cells.GetColumnWidth(i)}");
        }

        // Auto‑fit columns B through D (indexes 1‑3) to see overridden settings take effect
        sheet.AutoFitColumns(1, 3);

        // Display column widths after auto‑fit
        for (int i = 0; i <= 3; i++)
        {
            Console.WriteLine($"Column {(char)('A' + i)} width after autofit: {cells.GetColumnWidth(i)}");
        }

        // Save the workbook
        workbook.Save("StandardWidthAutoFitDemo.xlsx");
    }
}
