// Title: C# Example: Auto‑Fit All Columns Then Freeze Panes with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate cells, call Worksheet.AutoFitColumns() to size every column to its content, and then apply Worksheet.FreezePanes() so the calculated widths stay fixed when scrolling. The file is saved as an Excel workbook.
// Keywords: Aspose.Cells C# | AutoFitColumns | FreezePanes | preserve column width | Excel column autosize | worksheet freeze panes .NET | dynamic column width | Excel export C# | Aspose.Cells example | auto fit before freeze
// Common Searches: auto fit columns before freeze panes Aspose.Cells | keep column widths after freezing rows C# | Aspose.Cells AutoFitColumns then FreezePanes sample | C# Excel column autosize and freeze header | how to preserve column width when freezing panes
// Developer Intent: Adjust column widths automatically before locking rows/columns so the layout remains unchanged during scrolling.
// Use Cases: Generating reports where column sizes adapt to data and the header row stays visible. | Building spreadsheet templates that maintain consistent column widths across devices after freezing panes. | Exporting data to Excel with dynamic column sizing and a frozen top row/first column for improved readability.
// AI Prompts: Write C# code using Aspose.Cells to auto‑fit every column and then freeze the first row and column, ensuring widths are retained. | Explain why AutoFitColumns must be called before FreezePanes in Aspose.Cells and show a complete example. | Create a reusable method that receives a Workbook and a freeze‑pane configuration, auto‑fits all columns on each worksheet, and applies the freeze settings.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, populate cells, call Worksheet.AutoFitColumns() to size every column to its content, and then apply Worksheet.FreezePanes() so the calculated widths stay fixed when scrolling. The file is saved as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data to demonstrate column width changes
        worksheet.Cells["A1"].PutValue("Short");
        worksheet.Cells["B1"].PutValue("This is a considerably longer piece of text that should cause the column to expand");
        worksheet.Cells["C1"].PutValue("Medium length");
        worksheet.Cells["A2"].PutValue("Another short");
        worksheet.Cells["B2"].PutValue("Another very long text entry that will affect column B width");
        worksheet.Cells["C2"].PutValue("Text");

        // Auto‑fit all columns before freezing panes (feature rule: AutoFitColumns)
        worksheet.AutoFitColumns();

        // Freeze panes at cell B2 (row index 1, column index 1) with 1 frozen row and 1 frozen column
        // (feature rule: FreezePanes)
        worksheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("AutoFitAndFreezeDemo.xlsx");
    }
}
