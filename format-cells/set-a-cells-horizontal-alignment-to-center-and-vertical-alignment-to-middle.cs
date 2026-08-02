// Title: Aspose.Cells .NET: Center Text Horizontally & Vertically in a Cell
// Description: Shows how to create a workbook, put text into cell A1, build a Style with HorizontalAlignment and VerticalAlignment set to TextAlignmentType.Center, apply the style, and save the file as AlignedCell.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells cell alignment | C# center text Excel | horizontal alignment Aspose.Cells | vertical alignment Aspose.Cells | Excel style center .NET | TextAlignmentType.Center
// Common Searches: Aspose.Cells .NET center text in a cell | Set horizontal and vertical alignment with Aspose.Cells C# | How to align cell content to middle using Aspose.Cells | Apply centered style to Excel cell Aspose.Cells
// Developer Intent: The developer wants to align a cell’s content both horizontally and vertically (center‑middle).
// Use Cases: Formatting report headers so titles appear centered for readability. | Generating invoices where total amounts are centered within their cells. | Designing dashboard worksheets with labels that need both horizontal and vertical centering.
// AI Prompts: Generate C# code to apply bold font and center alignment to a range of cells with Aspose.Cells. | Provide examples of left, right, top, and bottom alignment settings for cells in Aspose.Cells .NET. | Explain how to create a reusable centered style and apply it across multiple worksheets using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsAlignmentDemo
{
    // Shows how to create a workbook, put text into cell A1, build a Style with HorizontalAlignment and VerticalAlignment set to TextAlignmentType.Center, apply the style, and save the file as AlignedCell.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Access cell A1 and put a sample value
            Cell cell = cells["A1"];
            cell.PutValue("Centered Text");

            // Create a style and set horizontal and vertical alignment to Center
            Style style = workbook.CreateStyle();
            style.HorizontalAlignment = TextAlignmentType.Center; // Horizontal center
            style.VerticalAlignment = TextAlignmentType.Center;   // Vertical middle

            // Apply the style to the cell
            cell.SetStyle(style);

            // Save the workbook to a file
            workbook.Save("AlignedCell.xlsx");
        }
    }
}
