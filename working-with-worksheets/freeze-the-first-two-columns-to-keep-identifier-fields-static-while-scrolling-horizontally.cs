// Title: How to Freeze the First Two Columns in an Excel Worksheet using Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, fills columns A and B with identifier values, adds sample data to columns C‑J, and calls worksheet.FreezePanes(0,2,0,2) to lock the first two columns while keeping all rows scrollable. The file is saved as FreezeFirstTwoColumns.xlsx, illustrating how to keep identifier fields static during horizontal scrolling.
// Keywords: Aspose.Cells | C# | .NET | FreezePanes | freeze first two columns | Excel column freeze | static identifier columns | horizontal scroll | worksheet.FreezePanes example | Excel automation
// Common Searches: Aspose.Cells freeze first two columns C# | FreezePanes method example .NET | How to lock columns A and B in Excel using Aspose.Cells | C# code to freeze columns in Excel workbook | Aspose.Cells keep identifier columns static while scrolling
// Developer Intent: Lock columns A and B so they remain visible when the user scrolls horizontally across the worksheet.
// Use Cases: Financial reports where ID and Account columns must stay visible while reviewing transaction data | Inventory sheets with SKU and Product Code frozen for easy reference | Data‑entry templates that protect identifier fields from scrolling out of view | Large datasets where row navigation is needed but key columns should stay static
// AI Prompts: Generate C# code to freeze the first three rows and first two columns in an Excel file using Aspose.Cells. | Explain each parameter of the FreezePanes method and how to calculate them for different freeze configurations. | Show how to programmatically unfreeze panes and then apply a new freeze setting with Aspose.Cells. | Provide a step‑by‑step guide to freeze columns based on a dynamic column index in a .NET application.

using System;
using Aspose.Cells;

namespace FreezeFirstTwoColumnsDemo
{
    // This C# example creates a new Workbook, fills columns A and B with identifier values, adds sample data to columns C‑J, and calls worksheet.FreezePanes(0,2,0,2) to lock the first two columns while keeping all rows scrollable. The file is saved as FreezeFirstTwoColumns.xlsx, illustrating how to keep identifier fields static during horizontal scrolling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data to visualize the freeze effect
            for (int row = 0; row < 20; row++)
            {
                // Identifier columns (to be frozen)
                worksheet.Cells[row, 0].PutValue($"ID{row + 1}");
                worksheet.Cells[row, 1].PutValue($"Code{row + 1}");

                // Additional data columns (scrollable)
                for (int col = 2; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first two columns (A and B)
            // Freeze at cell C1 (row index 0, column index 2) with 0 frozen rows and 2 frozen columns
            worksheet.FreezePanes(0, 2, 0, 2);

            // Save the workbook
            workbook.Save("FreezeFirstTwoColumns.xlsx");
        }
    }
}
