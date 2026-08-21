// Title: C# – Create Workbook, Merge A1:D1, Center Bold Title, Save with Aspose.Cells
// Description: Shows how to create a new workbook, merge cells A1:D1, insert a title, apply centered bold formatting, and save the file as MergedTitle.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | merge cells A1:D1 | center text | bold title | Excel workbook | save xlsx | cell alignment
// Common Searches: Aspose.Cells merge cells and center text C# | How to set a bold title across columns A to D with Aspose.Cells | Create merged title row in Excel using Aspose.Cells .NET | Set horizontal and vertical alignment for merged cells Aspose
// Developer Intent: Generate an Excel file with a merged title row that spans columns A‑D, centered and bold, then save the workbook.
// Use Cases: Standard report header with a spanning title | Invoice or receipt template that needs a prominent top title | Dashboard sheet where the main heading covers multiple columns | Automated data export that adds a formatted heading to each workbook
// AI Prompts: Generate C# Aspose.Cells code to merge A1:D1, set "Report Title" centered and bold, and save as MergedTitle.xlsx. | Explain step‑by‑step how to apply horizontal and vertical alignment to a merged cell in Aspose.Cells. | Show how to create a reusable style for merged title rows and apply it across multiple worksheets in C#.

using System;
using Aspose.Cells;

// Shows how to create a new workbook, merge cells A1:D1, insert a title, apply centered bold formatting, and save the file as MergedTitle.xlsx using Aspose.Cells for .NET.
class MergeTitleExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
        cells.Merge(0, 0, 1, 4);

        // Set the title text in the merged cell (upper‑left cell of the range)
        cells[0, 0].PutValue("Report Title");

        // Retrieve the style of the merged cell and set horizontal alignment to Center
        Style style = cells[0, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true; // optional: make the title bold
        cells[0, 0].SetStyle(style);

        // Save the workbook
        workbook.Save("MergedTitle.xlsx");
    }
}
