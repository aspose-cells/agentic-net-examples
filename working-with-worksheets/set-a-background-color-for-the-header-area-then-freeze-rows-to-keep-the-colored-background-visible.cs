// Title: Set Header Row Background Color and Freeze It with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, style the first row with a solid LightBlue fill and bold text using StyleFlag, lock the top row with FreezePanes, and save the result as HeaderFreezeDemo.xlsx via Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# header background | freeze panes Aspose.Cells | apply row style Aspose.Cells | Excel header color C# | StyleFlag background Aspose | worksheet FreezePanes example | Aspose.Cells .NET tutorial | Excel styling with Aspose | solid background color Aspose.Cells | freeze top row Excel
// Common Searches: How to color a header row in Excel using Aspose.Cells C# | Aspose.Cells freeze first row after applying style | Set background color for worksheet header with Aspose.Cells | C# code to style first row and freeze panes | Aspose.Cells example for header formatting and freeze panes
// Developer Intent: The developer wants to highlight the worksheet’s header row with a background shade and keep it visible while scrolling by freezing that row.
// Use Cases: Generating financial reports where column titles remain visible as users scroll through large data sets. | Exporting database tables to Excel with a colored, fixed header for easier navigation. | Building dashboard worksheets that require a styled top row locked in place for quick reference.
// AI Prompts: Create C# Aspose.Cells code that applies a gradient background to the first two header rows and freezes them. | Show how to set a custom font, background color, and border for multiple header rows and lock both rows and columns. | Explain the FreezePanes parameters for freezing rows, columns, and panes while preserving existing styles.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderFreezeDemo
{
    // Shows how to create a workbook, style the first row with a solid LightBlue fill and bold text using StyleFlag, lock the top row with FreezePanes, and save the result as HeaderFreezeDemo.xlsx via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a style for the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;               // Enable background fill
                headerStyle.BackgroundColor = Color.LightBlue;            // Set background color
                headerStyle.Font.IsBold = true;                           // Make header text bold

                // Define which style attributes to apply (background and font)
                StyleFlag flag = new StyleFlag();
                flag.CellShading = true;   // Apply background color
                flag.Font = true;          // Apply font settings

                // Apply the style to the first row (row index 0)
                cells.ApplyRowStyle(0, headerStyle, flag);

                // Freeze the first row so the colored header stays visible while scrolling
                // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook
                workbook.Save("HeaderFreezeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
