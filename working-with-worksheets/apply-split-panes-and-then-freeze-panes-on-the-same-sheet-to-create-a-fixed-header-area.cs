// Title: Split and Freeze Panes to Create a Fixed Header in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to split a worksheet window, then freeze the first five rows to keep a header visible while scrolling. The example populates sample data, calls Worksheet.Split(), applies Worksheet.FreezePanes(5,0,5,0), and saves the workbook as SplitAndFreezeDemo.xlsx.
// Keywords: Aspose.Cells split panes | Aspose.Cells freeze panes | fixed header Excel C# | Worksheet.Split | Worksheet.FreezePanes | C# Aspose.Cells example
// Common Searches: Aspose.Cells split panes and freeze rows | How to freeze top rows after splitting in Aspose.Cells .NET | Create fixed header in Excel using Aspose.Cells C# | Split window then freeze panes Aspose.Cells
// Developer Intent: The developer wants to split the worksheet view and then lock the first five rows so the header stays in place while the rest of the sheet scrolls.
// Use Cases: Build a large data report where column headers remain visible during vertical scrolling. | Design an interactive Excel dashboard that separates scrolling areas while keeping a static header. | Prepare a printable workbook where the top rows are locked for consistent header display across pages.
// AI Prompts: Show a C# snippet that splits a worksheet at row 5 and freezes the top five rows using Aspose.Cells. | Explain the difference between Worksheet.Split() and Worksheet.FreezePanes() and when to combine them. | Give guidance on freezing both rows and columns after a split with Aspose.Cells, including parameter calculations.

using System;
using Aspose.Cells;

// Demonstrates how to split a worksheet window, then freeze the first five rows to keep a header visible while scrolling. The example populates sample data, calls Worksheet.Split(), applies Worksheet.FreezePanes(5,0,5,0), and saves the workbook as SplitAndFreezeDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data (optional, just for visual verification)
        for (int row = 0; row < 30; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // First, split the window. This creates separate panes without freezing them.
        sheet.Split();

        // Then freeze the top rows to keep the header visible while scrolling.
        // Freeze the first 5 rows (row index 5) and no columns.
        sheet.FreezePanes(5, 0, 5, 0);

        // Save the workbook to a file
        workbook.Save("SplitAndFreezeDemo.xlsx");
    }
}
