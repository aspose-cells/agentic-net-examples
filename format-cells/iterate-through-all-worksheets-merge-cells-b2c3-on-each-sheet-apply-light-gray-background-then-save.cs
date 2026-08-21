// Title: C# – Merge B2:C3 on Every Worksheet and Apply Light Gray Fill with Aspose.Cells
// Description: Creates a workbook, adds optional sheets, defines a solid light‑gray style, iterates through all worksheets to merge the range B2:C3, applies the style to the merged area, and saves the file as MergedSheets.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | apply background color merged range | style merged cells Aspose.Cells | iterate worksheets format cells | save workbook Aspose.Cells
// Common Searches: merge same cell range on all worksheets Aspose.Cells | set light gray fill for merged cells C# | loop through worksheets and format merged range | Aspose.Cells example merge B2:C3 and style | save styled workbook with Aspose.Cells .NET
// Developer Intent: Merge cells B2:C3 on each worksheet, apply a light gray fill, and save the workbook.
// Use Cases: Create a uniform header spanning B2:C3 on every sheet with a gray background. | Standardize a title block across multiple worksheets in a financial report. | Prepare a template where a merged cell is consistently styled for corporate branding.
// AI Prompts: Write C# code using Aspose.Cells to merge B2:C3 on all worksheets and set a light gray background. | Show how to reuse a Style object for merged cells across multiple sheets in Aspose.Cells .NET. | Explain step‑by‑step how to iterate worksheets, merge a range, apply a solid fill, and save the workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    // Creates a workbook, adds optional sheets, defines a solid light‑gray style, iterates through all worksheets to merge the range B2:C3, applies the style to the merged area, and saves the file as MergedSheets.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (contains at least one worksheet)
            Workbook workbook = new Workbook();

            // Add a few extra worksheets for demonstration (optional)
            workbook.Worksheets.Add();
            workbook.Worksheets.Add();

            // Define the style with a light gray background
            Style grayStyle = workbook.CreateStyle();
            grayStyle.Pattern = BackgroundType.Solid;
            grayStyle.ForegroundColor = Color.LightGray;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Merge cells B2:C3 (zero‑based indices: row 1, column 1, 2 rows, 2 columns)
                sheet.Cells.Merge(1, 1, 2, 2);

                // Apply the gray background style to the merged cell (top‑left cell of the range)
                sheet.Cells[1, 1].SetStyle(grayStyle);
            }

            // Save the workbook to a file
            workbook.Save("MergedSheets.xlsx", SaveFormat.Xlsx);
        }
    }
}
