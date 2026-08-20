// Title: C# – Merge D4:F4 and Center Text Horizontally with Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells for .NET to merge the range D4:F4, insert a value, apply horizontal center alignment, and save the workbook as MergedAndCentered.xlsx.
// Keywords: Aspose.Cells | C# merge cells | merge D4 F4 | horizontal center alignment | Excel automation | worksheet formatting | Cells.Merge | TextAlignmentType.Center | save workbook | Aspose.Cells for .NET
// Common Searches: Aspose.Cells merge cells and center text C# | How to merge D4 to F4 in Aspose.Cells .NET | Set horizontal alignment after merging cells Aspose | C# code to merge Excel cells with Aspose.Cells
// Developer Intent: Merge the cells D4 through F4 into one range and align the contained text to the horizontal center using Aspose.Cells for .NET.
// Use Cases: Create a spanning title across columns D‑F in a generated report. | Design a centered header for an invoice or statement worksheet. | Place a merged label for a chart legend in an automated Excel file.
// AI Prompts: Write C# code that merges D4:F4 and sets HorizontalAlignment to Center with Aspose.Cells. | Show how to add bold and font‑size styling to the merged cell while keeping the text centered. | Explain how to retrieve and modify the style of a merged range after using Cells.Merge in Aspose.Cells. | Provide a step‑by‑step guide for merging cells and applying additional formatting in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace MergeAndCenterExample
{
    // Demonstrates how to use Aspose.Cells for .NET to merge the range D4:F4, insert a value, apply horizontal center alignment, and save the workbook as MergedAndCentered.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells D4 (row 3, column 3) through F4 (row 3, column 5)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns
            cells.Merge(3, 3, 1, 3);

            // Put a sample value into the merged cell (upper‑left cell of the range)
            cells[3, 3].PutValue("Merged and Centered");

            // Retrieve the style of the merged cell, set horizontal alignment to Center
            Style style = cells[3, 3].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            cells[3, 3].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("MergedAndCentered.xlsx");
        }
    }
}
