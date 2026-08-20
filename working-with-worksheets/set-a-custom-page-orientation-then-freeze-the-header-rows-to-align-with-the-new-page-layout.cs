// Title: Set Landscape Page Orientation and Freeze Header Row with Aspose.Cells for .NET
// Description: Creates a new workbook, switches the first worksheet to landscape mode, freezes the top row, repeats that header on every printed page, adds sample data, and saves the file as CustomOrientation_FreezeHeader.xlsx.
// Keywords: Aspose.Cells | C# | .NET | page orientation | landscape orientation | freeze panes | freeze header row | print title rows | worksheet layout | Excel export
// Common Searches: Aspose.Cells set landscape orientation C# | How to freeze the first row in Aspose.Cells | Repeat header row on each printed page Aspose.Cells | Freeze panes and set page orientation Aspose.Cells .NET
// Developer Intent: Apply a landscape orientation to a worksheet, freeze the first row, and ensure the header repeats on every printed page.
// Use Cases: Printable reports where the column headings must stay visible while scrolling and appear on each page. | Presentation‑ready spreadsheets that require a landscape layout with a persistent header for quick reference. | Invoices or statements that span multiple pages and need the title row repeated on every sheet.
// AI Prompts: Generate Aspose.Cells code to set portrait orientation, freeze the first two rows, and repeat them on printed pages. | Show how to configure custom margins, page orientation, and repeat both title rows and columns using Aspose.Cells for .NET. | Explain the FreezePanes parameters in Aspose.Cells with examples for freezing rows, columns, or a combination of both.

using System;
using Aspose.Cells;

// Creates a new workbook, switches the first worksheet to landscape mode, freezes the top row, repeats that header on every printed page, adds sample data, and saves the file as CustomOrientation_FreezeHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Set a custom page orientation (Landscape)
        // -------------------------------------------------
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // -------------------------------------------------
        // Freeze the header row (first row) so it stays visible
        // -------------------------------------------------
        // Freeze panes at cell A2: this freezes the first row (row index 1)
        worksheet.FreezePanes("A2", 1, 0);

        // Also repeat the header row on each printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // -------------------------------------------------
        // Add some sample data to demonstrate the effect
        // -------------------------------------------------
        worksheet.Cells["A1"].PutValue("Header");
        for (int i = 2; i <= 30; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Data row {i - 1}");
        }

        // Save the workbook
        workbook.Save("CustomOrientation_FreezeHeader.xlsx");
    }
}
