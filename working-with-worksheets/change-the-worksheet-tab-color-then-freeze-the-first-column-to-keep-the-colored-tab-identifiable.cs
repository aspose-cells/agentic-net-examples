// Title: Color Worksheet Tab and Freeze First Column with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to set a custom tab color for a worksheet and freeze column A using Aspose.Cells in C#. The example creates a workbook, applies a LightGreen tab color, freezes the first column with FreezePanes, and saves the file as TabColorAndFreezeFirstColumn.xlsx.
// Keywords: Aspose.Cells | C# | .NET | worksheet tab color | set TabColor | FreezePanes | freeze first column | Excel workbook example | change Excel tab color programmatically | freeze panes Aspose.Cells
// Common Searches: Aspose.Cells set worksheet tab color C# | How to freeze column A with Aspose.Cells | Freeze first column and change tab color in Excel using .NET | Aspose.Cells FreezePanes example | C# code to color Excel sheet tab and freeze panes
// Developer Intent: Apply a custom color to a worksheet tab and lock the first column so the colored tab stays identifiable while scrolling.
// Use Cases: Mark a summary or status sheet with a colored tab and keep identifier column visible during horizontal scrolling. | Create a multi‑sheet report where each sheet’s tab color reflects its category and the first column remains fixed for easy navigation. | Design a template that highlights key sheets by tab color while freezing the ID column for data entry consistency.
// AI Prompts: Generate C# code that sets a custom RGB tab color for multiple worksheets and freezes the first two columns using Aspose.Cells. | Show how to apply different tab colors based on worksheet names and freeze column A on each sheet with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsTabColorFreeze
{
    // Demonstrates how to set a custom tab color for a worksheet and freeze column A using Aspose.Cells in C#. The example creates a workbook, applies a LightGreen tab color, freezes the first column with FreezePanes, and saves the file as TabColorAndFreezeFirstColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the worksheet tab color (e.g., LightGreen)
            sheet.TabColor = Color.LightGreen;

            // Freeze the first column (column A)
            // Freeze at column index 1 (B) with 0 frozen rows and 1 frozen column
            sheet.FreezePanes(0, 1, 0, 1);

            // Save the workbook
            workbook.Save("TabColorAndFreezeFirstColumn.xlsx", SaveFormat.Xlsx);
        }
    }
}
