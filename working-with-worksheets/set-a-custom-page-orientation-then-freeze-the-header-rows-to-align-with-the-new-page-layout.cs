// Title: Aspose.Cells for .NET: Landscape Orientation + Freeze Header Row with Print Title
// Description: Shows how to create a workbook, add a header row, set the worksheet to landscape orientation, freeze the first row, repeat that row on every printed page, and save the file as XLSX using C#.
// Keywords: Aspose.Cells C# page orientation | Aspose.Cells freeze panes | Aspose.Cells repeat print titles | landscape orientation Aspose.Cells | freeze header row C# | print title rows Aspose.Cells | custom page layout .NET
// Common Searches: C# set worksheet to landscape Aspose.Cells | freeze top row Aspose.Cells .NET | repeat header row on each printed page Aspose.Cells | how to freeze panes and set print titles in C# | Aspose.Cells page setup orientation and freeze panes example
// Developer Intent: The developer wants the worksheet printed in landscape mode, the header row locked while scrolling, and the same header to appear on each printed page.
// Use Cases: Generating wide reports that need landscape printing while keeping column headers visible during scrolling. | Creating invoices or schedules where the top row must stay fixed and repeat on every printed sheet. | Building dashboards exported to Excel with a consistent header across all pages.
// AI Prompts: Provide C# code using Aspose.Cells to set portrait orientation, freeze the first three rows, and set them as print titles. | Explain how to adjust column widths after freezing panes so the header aligns with data columns. | Write a method that switches page orientation between portrait and landscape based on column count.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add a header row, set the worksheet to landscape orientation, freeze the first row, repeat that row on every printed page, and save the file as XLSX using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data with a header row
        worksheet.Cells["A1"].PutValue("Header");
        for (int i = 2; i <= 30; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Data row {i - 1}");
        }

        // Set custom page orientation (Landscape)
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Freeze the header row so it stays visible while scrolling
        // Freeze first row (row index 1, column index 0) with 1 frozen row and 0 frozen columns
        worksheet.FreezePanes(1, 0, 1, 0);

        // Ensure the header row repeats on each printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Save the workbook (save rule)
        workbook.Save("CustomOrientationAndFreeze.xlsx", SaveFormat.Xlsx);
    }
}
