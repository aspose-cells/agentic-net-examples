// Title: Color Worksheet Tab and Freeze First Column with Aspose.Cells for .NET (C#)
// Description: Shows how to assign a custom color to a worksheet tab and freeze column A using Aspose.Cells for .NET, then save the file as TabColorAndFreezeFirstColumn.xlsx.
// Keywords: Aspose.Cells | C# | worksheet tab color | freeze first column | FreezePanes | Excel tab color .NET | column A freeze | set tab color programmatically | Excel UI customization
// Common Searches: Aspose.Cells change worksheet tab color C# | Freeze first column Aspose.Cells .NET | How to lock column A with FreezePanes in Excel using Aspose | Set tab color programmatically with Aspose.Cells | C# example for coloring Excel sheet tab and freezing a column
// Developer Intent: Apply a distinct tab color to a worksheet and keep column A fixed during scrolling.
// Use Cases: Mark a dashboard sheet with a bright tab and freeze the identifier column for quick reference while scrolling through data. | Create department-specific worksheets where each tab has a unique color and the first column remains visible for row labels. | Generate export files that combine visual navigation (colored tabs) with static key columns to improve user experience in large reports.
// AI Prompts: Provide a C# snippet that sets a worksheet tab to a custom RGB value and freezes the first two columns using Aspose.Cells. | Explain how to conditionally change the tab color based on sheet content and apply FreezePanes to keep the header column static. | Show examples of FreezePanes with different cell references to lock rows, columns, or both in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to assign a custom color to a worksheet tab and freeze column A using Aspose.Cells for .NET, then save the file as TabColorAndFreezeFirstColumn.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet tab color (e.g., blue) to make it easily identifiable
            worksheet.TabColor = Color.Blue;

            // Freeze the first column.
            // FreezePanes with cell "B1" freezes columns to the left of column B (i.e., column A)
            // No rows are frozen (0), and 1 column is frozen.
            worksheet.FreezePanes("B1", 0, 1);

            // Save the workbook
            workbook.Save("TabColorAndFreezeFirstColumn.xlsx");
        }
    }
}
