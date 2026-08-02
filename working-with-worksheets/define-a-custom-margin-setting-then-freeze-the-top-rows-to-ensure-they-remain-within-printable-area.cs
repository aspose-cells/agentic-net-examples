// Title: Set custom page margins (cm) and freeze the top row using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, applies 2 cm top, 1.5 cm bottom, and 1 cm side margins, repeats the first row on every printed page, freezes that row at A2, and saves the file as CustomMarginAndFreeze.xlsx.
// Keywords: Aspose.Cells page margins centimeters | freeze top row C# | PrintTitleRows Aspose.Cells | custom margin .NET | FreezePanes Aspose.Cells | Aspose.Cells printing settings | C# Excel page setup
// Common Searches: Aspose.Cells set page margins in centimeters | How to freeze the first row in Aspose.Cells .NET | Repeat header row on each printed page Aspose.Cells | Combine custom margins and freeze panes Aspose.Cells | C# code for page setup and FreezePanes
// Developer Intent: Apply specific page margins and keep the header row visible and printable by freezing it.
// Use Cases: Generate a printable report with precise top, bottom, and side margins while keeping the header row on every page. | Create a large data sheet where the first row stays fixed during scrolling and is included in the print layout. | Prepare Excel files for corporate printing standards that require custom margins and frozen header rows.
// AI Prompts: Write C# code with Aspose.Cells to set top, bottom, left, and right margins in centimeters and freeze the first row. | Show how to use PrintTitleRows to repeat a header row on each printed page and freeze that row at A2. | Explain the steps to combine custom page margins and FreezePanes so the header remains within the printable area.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample data, applies 2 cm top, 1.5 cm bottom, and 1 cm side margins, repeats the first row on every printed page, freezes that row at A2, and saves the file as CustomMarginAndFreeze.xlsx.
class CustomMarginAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Header");
        for (int i = 2; i <= 20; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
        }

        // Configure page margins (centimeters)
        PageSetup pageSetup = sheet.PageSetup;
        pageSetup.TopMargin = 2.0;      // Top margin = 2 cm
        pageSetup.BottomMargin = 1.5;   // Bottom margin = 1.5 cm
        pageSetup.LeftMargin = 1.0;     // Left margin = 1 cm
        pageSetup.RightMargin = 1.0;    // Right margin = 1 cm

        // Repeat the first row on every printed page
        pageSetup.PrintTitleRows = "$1:$1";

        // Freeze the top row so it stays visible while scrolling
        // Freeze at cell A2, freezing 1 row and 0 columns
        sheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("CustomMarginAndFreeze.xlsx");
    }
}
