// Title: How to Set an Excel Worksheet Tab Color to Teal with Aspose.Cells for .NET (C#)
// Description: This example demonstrates creating a workbook with Aspose.Cells, accessing the first worksheet, applying the System.Drawing.Color.Teal value to the TabColor property, and saving the file as TabColorTeal.xlsx, resulting in a teal‑colored sheet tab.
// Keywords: Aspose.Cells | C# Excel | worksheet tab color | set tab color | teal | .NET | Excel sheet tab color | Aspose.Cells TabColor
// Common Searches: Aspose.Cells change worksheet tab color C# | set Excel sheet tab to teal using .NET | C# code for worksheet TabColor property | how to color Excel sheet tabs with Aspose.Cells | example of TabColor teal Aspose.Cells
// Developer Intent: Apply a teal color to a worksheet tab in an Excel file using Aspose.Cells.
// Use Cases: Mark a dashboard sheet with a distinctive teal tab for quick identification. | Apply corporate branding by setting all worksheet tabs to the brand's teal shade. | Highlight newly generated reports by assigning them a teal tab color.
// AI Prompts: Generate C# code with Aspose.Cells that sets the second worksheet's tab to teal and saves the workbook as 'Report.xlsx'. | Show how to loop through worksheets and change each tab to teal when a specific condition (e.g., sheet name contains 'Summary') is met.

using System;
using System.Drawing;
using Aspose.Cells;

// This example demonstrates creating a workbook with Aspose.Cells, accessing the first worksheet, applying the System.Drawing.Color.Teal value to the TabColor property, and saving the file as TabColorTeal.xlsx, resulting in a teal‑colored sheet tab.
class SetWorksheetTabColor
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet tab color to teal
        worksheet.TabColor = Color.Teal;

        // Save the workbook
        workbook.Save("TabColorTeal.xlsx");
    }
}
