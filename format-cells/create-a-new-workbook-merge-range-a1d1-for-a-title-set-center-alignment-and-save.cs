// Title: C# – Create a Workbook, Merge A1:D1 as a Centered Title, and Save with Aspose.Cells
// Description: Demonstrates how to instantiate a new Workbook, merge the range A1:D1 on the first worksheet, insert a title, apply horizontal and vertical center alignment, and save the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | center align merged cell Aspose | create workbook with title row | Excel merge A1:D1 Aspose.Cells | set TextAlignmentType Center | save workbook Aspose.Cells .NET | style merged cell Aspose | C# Excel automation Aspose.Cells
// Common Searches: how to merge cells A1 to D1 using Aspose.Cells C# | center text in merged cells Aspose.Cells .NET | save Excel file after merging cells with Aspose | apply vertical and horizontal alignment to merged cell C# | Aspose.Cells example for title row merge
// Developer Intent: Generate a new Excel workbook, merge the first row for a title, center the title text, and write the file to disk.
// Use Cases: Create a report header where the title spans columns A‑D and is centrally aligned before adding data. | Design an invoice template with a merged, centered heading across the top of the sheet. | Build a dashboard worksheet that requires a bold, centered title row prior to inserting charts and tables.
// AI Prompts: Write C# code with Aspose.Cells to merge cells A1:D1, set a custom title, apply both horizontal and vertical center alignment, and save as 'Report.xlsx'. | Explain how to modify the style of a merged cell in Aspose.Cells for .NET to include font size, bold formatting, and centered alignment. | Provide an example that adds a merged title row to an existing workbook while preserving the original formatting of other cells.

using System;
using Aspose.Cells;

namespace AsposeCellsMergeTitleExample
{
    // Demonstrates how to instantiate a new Workbook, merge the range A1:D1 on the first worksheet, insert a title, apply horizontal and vertical center alignment, and save the file as an Excel workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
            cells.Merge(0, 0, 1, 4);

            // Set the title text in the merged cell
            cells[0, 0].PutValue("Report Title");

            // Retrieve the style of the merged cell and set horizontal alignment to center
            Style style = cells[0, 0].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            // Optionally, set vertical alignment to center as well
            style.VerticalAlignment = TextAlignmentType.Center;
            // Apply the updated style back to the cell
            cells[0, 0].SetStyle(style);

            // Save the workbook to a file
            workbook.Save("MergedTitle.xlsx");
        }
    }
}
