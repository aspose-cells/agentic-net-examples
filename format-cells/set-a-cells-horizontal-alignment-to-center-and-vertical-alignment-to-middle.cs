// Title: Center text horizontally and vertically in an Excel cell using Aspose.Cells for .NET
// Description: C# sample that creates a workbook, writes "Centered Text" to cell A1, retrieves the cell's Style, sets HorizontalAlignment and VerticalAlignment to TextAlignmentType.Center, reapplies the style, and saves the file as CellAlignmentDemo.xlsx.
// Keywords: Aspose.Cells C# cell alignment | Excel cell horizontal center | Excel cell vertical middle | TextAlignmentType.Center Aspose | programmatic cell styling Aspose.Cells | centered cell content .NET | format cell alignment Aspose
// Common Searches: Aspose.Cells set cell alignment C# | center text in Excel cell using Aspose.Cells | horizontal and vertical alignment Aspose.Cells .NET | C# code to center cell content in workbook | apply centered style to Excel cell with Aspose
// Developer Intent: Apply centered horizontal and vertical alignment to a specific cell in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a report header where the title spans a merged range and appears perfectly centered. | Design a dashboard sheet with label cells that need both axes centered for visual balance. | Generate invoices where the company name cell is aligned to the middle for a professional layout.
// AI Prompts: Generate C# code with Aspose.Cells that centers a cell's text horizontally and vertically and saves the workbook. | Explain how to modify a cell's Style to use TextAlignmentType.Center for both HorizontalAlignment and VerticalAlignment in Aspose.Cells. | Show an example of applying centered alignment to an entire range of cells using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// C# sample that creates a workbook, writes "Centered Text" to cell A1, retrieves the cell's Style, sets HorizontalAlignment and VerticalAlignment to TextAlignmentType.Center, reapplies the style, and saves the file as CellAlignmentDemo.xlsx.
class SetCellAlignment
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1 and set a value
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Centered Text");

        // Retrieve the cell's style
        Style style = cell.GetStyle();

        // Set horizontal and vertical alignment to Center
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the modified style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("CellAlignmentDemo.xlsx");
    }
}
