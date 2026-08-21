// Title: Merge cells, apply bold font, and center text in a header row with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, merges cells A1‑D1 on the first worksheet, inserts a header title, sets the font to bold, aligns the text horizontally to the center, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | header row formatting Aspose.Cells | bold font Aspose.Cells | center alignment merged cells | save workbook Aspose.Cells .NET
// Common Searches: Aspose.Cells merge cells and center text C# | how to make header bold and centered with Aspose.Cells | C# code for merged header row Aspose.Cells
// Developer Intent: Create a merged header row, make its text bold and horizontally centered, then save the workbook.
// Use Cases: Generating a report with a bold, centered title that spans multiple columns. | Designing an invoice template where the company name header merges across columns and uses consistent styling. | Building a dashboard worksheet with section titles formatted as merged, bold, centered headers.
// AI Prompts: Show C# code to merge cells A1 to D1, set the text bold, and center it using Aspose.Cells. | Provide an Aspose.Cells example that formats a merged header cell with larger font size, bold weight, and center alignment. | Explain how to create a reusable style for multiple merged header rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new workbook, merges cells A1‑D1 on the first worksheet, inserts a header title, sets the font to bold, aligns the text horizontally to the center, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells in the first row (A1:D1)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        cells.Merge(0, 0, 1, 4);

        // Set the header text in the merged cell
        cells[0, 0].PutValue("Header Title");

        // Retrieve the style of the merged cell
        Style style = cells[0, 0].GetStyle();

        // Apply bold font
        style.Font.IsBold = true;

        // Center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the modified style back to the merged cell
        cells[0, 0].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedHeader.xlsx");
    }
}
