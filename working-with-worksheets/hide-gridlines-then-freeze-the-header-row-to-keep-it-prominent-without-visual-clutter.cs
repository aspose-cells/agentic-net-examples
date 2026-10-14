// Title: Hide gridlines and freeze the header row in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, hides the worksheet gridlines, and freezes the first row. | Demonstrate how to call the FreezePanes method to keep the header row visible while disabling gridlines in an Excel file. | Provide a complete example that saves the modified workbook as Output.xlsx after applying gridline hiding and freeze‑pane settings.
// Common Searches: Aspose.Cells C# hide gridlines and freeze top row in Excel | how to use FreezePanes to lock header row while turning off gridlines with Aspose.Cells | C# example for removing worksheet gridlines and freezing first row using Aspose.Cells library
// Tags: hide worksheet gridlines Aspose.Cells C# | freeze header row FreezePanes Aspose.Cells | Aspose.Cells visual cleanup Excel workbook | disable gridlines set freeze panes C# | Excel output file without gridlines Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a new workbook, hides gridlines on the first worksheet, freezes the first row with FreezePanes, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Hide gridlines to reduce visual clutter
        sheet.IsGridlinesVisible = false;

        // Freeze the header row (first row) so it stays visible while scrolling
        // Parameters: totalRows to freeze, totalColumns to freeze, first scrollable row, first scrollable column
        sheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
