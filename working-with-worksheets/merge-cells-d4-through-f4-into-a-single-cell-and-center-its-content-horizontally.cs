// Title: Merge D4‑F4 and center text horizontally with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, merges cells D4 through F4 on the first worksheet, optionally writes a value, applies a style that sets HorizontalAlignment to Center, and saves the file as MergedCentered.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | horizontal alignment Aspose.Cells | merge range D4 F4 | center text merged cell | C# Excel cell merging | Aspose.Cells style settings
// Common Searches: Aspose.Cells merge cells D4 to F4 C# | how to center text in merged cells using Aspose.Cells | C# merge cells and set horizontal alignment Aspose.Cells | Aspose.Cells merge range and apply style | centered header across columns D-F Aspose.Cells
// Developer Intent: The developer wants to merge cells D4 through F4 into a single cell and center its content horizontally.
// Use Cases: Generate a report header that spans columns D‑F with a centered title. | Create a summary label by merging adjacent cells and aligning the text for clear presentation. | Design a template where section headings occupy merged cells and are horizontally centered for visual consistency.
// AI Prompts: Write C# code with Aspose.Cells to merge cells A1:C1, make the text bold, and center it both horizontally and vertically. | Show how to apply horizontal and vertical alignment to a merged cell range using Aspose.Cells for .NET. | Explain how to merge cells without losing existing data and then apply a custom style to the merged cell in Aspose.Cells.

using Aspose.Cells;

// Creates a new workbook, merges cells D4 through F4 on the first worksheet, optionally writes a value, applies a style that sets HorizontalAlignment to Center, and saves the file as MergedCentered.xlsx using Aspose.Cells for .NET.
class MergeAndCenterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells D4 through F4 (zero‑based indices: row 3, column 3, 1 row, 3 columns)
        worksheet.Cells.Merge(3, 3, 1, 3);

        // Optional: put a value into the merged cell
        worksheet.Cells[3, 3].PutValue("Merged and Centered");

        // Center the content horizontally
        Style style = worksheet.Cells[3, 3].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        worksheet.Cells[3, 3].SetStyle(style);

        // Save the workbook
        workbook.Save("MergedCentered.xlsx");
    }
}
