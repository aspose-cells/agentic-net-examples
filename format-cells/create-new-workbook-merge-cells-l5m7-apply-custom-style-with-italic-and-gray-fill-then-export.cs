// Title: C# – Merge L5:M7 and apply italic gray style using Aspere.Cells
// Description: Creates a new Workbook, merges the range L5:M7 (rows 5‑7, columns L‑M), defines a custom Style with an italic font and solid gray fill, applies it to the merged block, and saves the result as MergedStyled.xlsx.
// Keywords: Aspose.Cells C# | merge cells range | custom style italic | gray background fill | Workbook export XLSX | merged cell formatting | L5:M7 styling | .NET spreadsheet library
// Common Searches: how to merge a cell range and style it with Aspose.Cells .NET | apply italic font and gray background to merged cells in C# | Aspose.Cells example for merging L5:M7 | create custom style for merged cells using Aspose.Cells | save workbook with styled merged range in C#
// Developer Intent: Produce a workbook where L5:M7 is merged and displayed with italic gray formatting.
// Use Cases: Header section that spans multiple rows/columns with emphasized italic‑gray appearance. | Template title blocks where merged cells need a distinct style for readability. | Printable forms or invoices that require a shaded, italicized merged area for section labels.
// AI Prompts: Generate C# code that merges cells L5:M7, sets an italic font with a solid gray fill, and saves the workbook using Aspose.Cells. | Show an Aspose.Cells .NET snippet for creating a custom style, applying it to a merged range, and exporting to XLSX.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergeStyleDemo
{
    // Creates a new Workbook, merges the range L5:M7 (rows 5‑7, columns L‑M), defines a custom Style with an italic font and solid gray fill, applies it to the merged block, and saves the result as MergedStyled.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells L5:M7 (rows 5-7, columns L-M)
            // Zero‑based indices: row 4, column 11, 3 rows, 2 columns
            worksheet.Cells.Merge(4, 11, 3, 2);

            // Create a custom style using the Workbook.CreateStyle method
            Style customStyle = workbook.CreateStyle();
            // Set italic font
            customStyle.Font.IsItalic = true;
            // Set gray fill
            customStyle.ForegroundColor = Color.Gray;
            customStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the merged cell (upper‑left cell of the range)
            worksheet.Cells[4, 11].SetStyle(customStyle);

            // Save the workbook
            workbook.Save("MergedStyled.xlsx");
        }
    }
}
