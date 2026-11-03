// Title: Insert a hyperlink with custom display text into a specific Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: C# how to add a hyperlink with display text to a cell using Aspose.Cells | Aspose.Cells Hyperlinks.Add example for setting URL and visible text | programmatically create a clickable link in an Excel worksheet with Aspose.Cells C# | set cell B2 to show 'Visit Aspose' and link to https://www.aspose.com using Aspose.Cells | add a single‑cell hyperlink in a .NET workbook with custom label
// Tags: Aspose.Cells Hyperlinks.Add method | C# insert Excel cell hyperlink | custom display text for Excel hyperlink Aspose | single‑cell URL link .NET workbook | save workbook with hyperlink Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, and uses the Hyperlinks.Add method to place a URL in cell B2. It then assigns custom display text to the cell so the link appears as "Visit Aspose". Finally, the workbook is saved as HyperlinkDemo.xlsx, with exception handling for robustness.
class HyperlinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell address where the hyperlink will be placed
            string cellAddress = "B2";

            // Define the display text and the URL
            string displayText = "Visit Aspose";
            string url = "https://www.aspose.com";

            // Get the target cell and its zero‑based row/column indices
            Cell targetCell = sheet.Cells[cellAddress];
            int row = targetCell.Row;
            int column = targetCell.Column;

            // Add the hyperlink to the specified cell (1 row × 1 column range)
            // Overload: Add(firstRow, firstColumn, totalRows, totalColumns, url)
            sheet.Hyperlinks.Add(row, column, 1, 1, url);

            // Set the cell value to the display text (ensures visible text)
            targetCell.PutValue(displayText);

            // Save the workbook
            workbook.Save("HyperlinkDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
