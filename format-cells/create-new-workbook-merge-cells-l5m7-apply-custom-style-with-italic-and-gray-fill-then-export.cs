// Title: Merge cells L5:M7 and apply an italic gray style with Aspose.Cells for .NET (C#)
// Description: C# code that creates a new Workbook, merges the range L5:M7 (rows 5‑7, columns L‑M), defines a custom style with an italic font and solid gray fill, applies the style to the merged range, and saves the file as MergedStyled.xlsx.
// Keywords: Aspose.Cells | C# | merge cells | L5:M7 | custom style | italic font | gray fill | solid background | Excel export | Workbook.Save
// Common Searches: Aspose.Cells merge cells C# example | apply custom style to merged cells Aspose.Cells | italic gray background Excel using Aspose.Cells | save workbook as XLSX with Aspose.Cells | how to merge range L5:M7 in Aspose.Cells
// Developer Intent: Merge L5:M7, set an italic gray style, and save the workbook.
// Use Cases: Create a multi‑row header for a financial report with emphasized formatting. | Design a highlighted title block in a dashboard by merging cells and applying a gray background. | Prepare a printable invoice section where the merged cells need italic text for stylistic emphasis.
// AI Prompts: Generate C# code using Aspose.Cells that merges cells B2:C4, applies a bold red font with a yellow background, and saves the workbook as Styled.xlsx. | Show how to define a custom style, apply it to any merged range, and export the result to PDF with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsMergeStyleDemo
{
    // C# code that creates a new Workbook, merges the range L5:M7 (rows 5‑7, columns L‑M), defines a custom style with an italic font and solid gray fill, applies the style to the merged range, and saves the file as MergedStyled.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells L5:M7 (zero‑based indices: row 4, column 11, 3 rows, 2 columns)
            worksheet.Cells.Merge(4, 11, 3, 2);

            // Create a custom style: italic font and gray fill
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsItalic = true;
            customStyle.ForegroundColor = Color.Gray;
            customStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the merged cell (upper‑left cell of the range)
            worksheet.Cells[4, 11].SetStyle(customStyle);

            // Save the workbook
            workbook.Save("MergedStyled.xlsx");
        }
    }
}
