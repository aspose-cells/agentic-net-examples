// Title: C# – Merge Z4:AA4 and Apply 45° Text Rotation with Aspose.Cells, Save as XLSX
// Description: Load an existing workbook, merge the range Z4:AA4 on the first worksheet, set a 45‑degree text rotation via a style and StyleFlag, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells | C# rotate text | 45 degree text rotation | Excel style flag | save workbook as xlsx | merged cell formatting | Aspose.Cells .NET example
// Common Searches: Aspose.Cells merge Z4 AA4 C# | rotate text 45 degrees merged cell Aspose | C# set cell rotation Aspose.Cells | how to merge cells and apply style Aspose | save modified workbook as xlsx Aspose.Cells
// Developer Intent: Merge cells Z4:AA4, rotate the text 45°, and export the workbook to XLSX with Aspose.Cells in C#.
// Use Cases: Create a diagonal header spanning two columns for financial dashboards. | Design printable forms where narrow columns need angled labels. | Build a template with merged title cells that stand out visually.
// AI Prompts: Write C# code using Aspose.Cells to merge Z4:AA4, set a 45° rotation, and save as output.xlsx. | Explain why a StyleFlag is required when applying rotation to a merged cell in Aspose.Cells. | Provide a step‑by‑step tutorial for modifying an existing XLSX file to merge cells and rotate text with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an existing workbook, merge the range Z4:AA4 on the first worksheet, set a 45‑degree text rotation via a style and StyleFlag, and save the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the existing spreadsheet
        string inputPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells Z4:AA4 (row index 3, column index 25, 1 row, 2 columns)
        worksheet.Cells.Merge(3, 25, 1, 2);

        // Create a style with a 45‑degree text rotation
        Style style = workbook.CreateStyle();
        style.RotationAngle = 45;

        // Enable the rotation flag so the style is applied
        StyleFlag flag = new StyleFlag();
        flag.Rotation = true;

        // Apply the style to the merged cell (upper‑left cell Z4)
        worksheet.Cells[3, 25].SetStyle(style, flag);

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx");
    }
}
