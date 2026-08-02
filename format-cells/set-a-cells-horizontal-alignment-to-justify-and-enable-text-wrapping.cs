// Title: Apply Justify Alignment and Text Wrapping to a Cell with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set a cell's horizontal alignment to Justify, enable text wrapping, auto‑fit the row height, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells justify alignment C# | Aspose.Cells text wrap | horizontal alignment justify Aspose.Cells | cell style wrap text .NET | auto fit row Aspose.Cells | C# Aspose.Cells formatting
// Common Searches: Aspose.Cells set cell justification C# | how to wrap text in a cell using Aspose.Cells .NET | justify alignment with text wrap Aspose.Cells example | C# code for justified and wrapped cell in Aspose.Cells
// Developer Intent: The developer needs to format a specific cell so its content is justified and automatically wrapped, with the row height adjusted to display all lines.
// Use Cases: Creating a report where paragraph text must be justified and wrapped inside a single cell. | Designing an invoice template that aligns product descriptions with justified text and wraps long strings. | Generating a data export where cells contain multi‑line notes that require both justification and automatic row height adjustment.
// AI Prompts: Show C# code to apply justify alignment and enable text wrapping for a range of cells in Aspose.Cells. | Explain how to set justify alignment and wrap text for merged cells using Aspose.Cells for .NET. | Provide a snippet that adjusts row height after enabling text wrap and justify alignment with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set a cell's horizontal alignment to Justify, enable text wrapping, auto‑fit the row height, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1 and put a long text value
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("This is a long text that should be justified and wrapped within the cell.");

        // Retrieve the cell's style, set justification and enable text wrapping
        Style style = cell.GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Justify; // justify alignment
        style.IsTextWrapped = true;                             // enable wrapping
        cell.SetStyle(style);                                   // apply the style to the cell

        // Adjust row height so the wrapped text is visible
        worksheet.AutoFitRow(0);

        // Save the workbook
        workbook.Save("JustifyWrap.xlsx");
    }
}
