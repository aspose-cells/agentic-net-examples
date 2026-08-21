// Title: C# – Merge A1:F2 into a Title Block with Aspose.Cells
// Description: Demonstrates how to create a new workbook, merge the range A1:F2, insert a centered bold title, and save the file as TitleBlock.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | Excel title block Aspose | merge A1 F2 Aspose.Cells | centered bold title Excel | Aspose.Cells styling merged range | C# Excel workbook example
// Common Searches: how to merge A1 to F2 with Aspose.Cells | set title in merged cells C# Aspose | center bold text in merged Excel range | Aspose.Cells example for title block | C# code to merge multiple rows and columns in Excel
// Developer Intent: Create a merged title block spanning A1‑F2, apply centered bold formatting, and export the workbook.
// Use Cases: Report header that covers the first two rows and six columns | Invoice or statement header spanning the top of a worksheet | Dashboard title area across the sheet’s upper rows
// AI Prompts: Provide C# code that merges cells A1:F2, adds "Report Title", centers the text, makes it bold, and saves the workbook with Aspose.Cells. | Show an Aspose.Cells .NET example for creating a multi‑row, multi‑column title block with styling. | How can I generate a merged title block covering A1 to F2 and export it as an Excel file using Aspose.Cells?

using System;
using Aspose.Cells;

namespace MergeTitleBlockDemo
{
    // Demonstrates how to create a new workbook, merge the range A1:F2, insert a centered bold title, and save the file as TitleBlock.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells from A1 (row 0, column 0) to F2 (row 1, column 5)
            // totalRows = 2 (rows 0 and 1), totalColumns = 6 (columns A to F)
            cells.Merge(0, 0, 2, 6);

            // Set the title text in the merged cell (upper‑left cell of the range)
            cells[0, 0].PutValue("Report Title");

            // Apply basic styling: center alignment and bold font
            Style style = cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Font.IsBold = true;
            cells[0, 0].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("TitleBlock.xlsx");
        }
    }
}
