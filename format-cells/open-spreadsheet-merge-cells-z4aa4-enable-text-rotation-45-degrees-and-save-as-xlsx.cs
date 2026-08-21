// Title: Merge Z4:AA4 and apply 45° text rotation with Aspose.Cells for .NET (C#)
// Description: Load an existing XLSX workbook, merge the range Z4:AA4 on the first worksheet, create a style with a 45‑degree rotation, apply it to the merged cell, and save the result as a new XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | Aspose.Cells text rotation | C# rotate text in Excel | Aspose.Cells style flag rotation | save workbook as xlsx Aspose
// Common Searches: Aspose.Cells merge Z4 AA4 C# | rotate text 45 degrees merged cells Aspose | how to apply style to merged cells Aspose.Cells | save modified Excel as XLSX using Aspose.Cells
// Developer Intent: Merge a specific cell range, rotate its text 45°, and save the workbook as XLSX with Aspose.Cells in C#.
// Use Cases: Create a diagonal header spanning two columns for a financial summary. | Design a dashboard title row with angled text for visual emphasis. | Prepare printable reports where merged cells need slanted labels for space efficiency.
// AI Prompts: Show C# code to merge Z4:AA4, set a 45° rotation, and save as XLSX with Aspose.Cells. | Explain how to use StyleFlag to enable text rotation on a merged cell in Aspose.Cells. | Provide a step‑by‑step example of applying a rotation style to merged cells and exporting the workbook.

using System;
using Aspose.Cells;

// Load an existing XLSX workbook, merge the range Z4:AA4 on the first worksheet, create a style with a 45‑degree rotation, apply it to the merged cell, and save the result as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        string inputPath = "input.xlsx";               // replace with your source file
        Workbook workbook = new Workbook(inputPath);

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells Z4:AA4
        // Row index is zero‑based (4th row -> 3), column Z is 25, AA is 26
        // Merge 1 row and 2 columns starting at (3,25)
        sheet.Cells.Merge(3, 25, 1, 2);

        // Create a style with a 45‑degree text rotation
        Style rotStyle = workbook.CreateStyle();
        rotStyle.RotationAngle = 45;

        // Enable the rotation flag so the style is applied
        StyleFlag flag = new StyleFlag();
        flag.Rotation = true;

        // Apply the style to the merged cell (upper‑left cell Z4)
        sheet.Cells["Z4"].SetStyle(rotStyle, flag);

        // Save the modified workbook as XLSX
        string outputPath = "output.xlsx";             // desired output file
        workbook.Save(outputPath);
    }
}
