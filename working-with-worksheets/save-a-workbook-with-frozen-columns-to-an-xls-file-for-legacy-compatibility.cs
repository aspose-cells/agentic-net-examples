// Title: Save a workbook with frozen columns to XLS using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate data, freeze the first two columns with FreezePanes, and save it as an XLS file using XlsSaveOptions for Excel 97‑2003 compatibility.
// Keywords: Aspose.Cells FreezePanes C# | save workbook as XLS | XlsSaveOptions legacy Excel | freeze columns Aspose.Cells | .NET export frozen panes | Excel 97-2003 compatibility Aspose
// Common Searches: Aspose.Cells freeze first two columns and save as xls | How to use FreezePanes with XlsSaveOptions in C# | Export workbook with frozen panes to legacy Excel format | C# code to freeze columns and save to .xls using Aspose.Cells
// Developer Intent: Generate an XLS file that retains frozen columns by applying FreezePanes and XlsSaveOptions in Aspose.Cells for .NET.
// Use Cases: Create reports where ID and Name columns stay visible while scrolling horizontally, then distribute them to users with older Excel versions. | Export data from a .NET application to a legacy system that only accepts .xls files, preserving pane freezing for readability. | Automate spreadsheet generation with frozen columns for dashboards that must be opened in Excel 97‑2003.
// AI Prompts: Show C# code to freeze the first three columns in an Aspose.Cells workbook and save it as an XLS file. | Explain how XlsSaveOptions works with FreezePanes to keep frozen panes when exporting to legacy Excel format. | Provide a step‑by‑step guide for saving a workbook with frozen columns to .xls using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, populate data, freeze the first two columns with FreezePanes, and save it as an XLS file using XlsSaveOptions for Excel 97‑2003 compatibility.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill some sample data
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Score");
        for (int i = 2; i <= 10; i++)
        {
            worksheet.Cells[i, 0].PutValue(i - 1);               // ID
            worksheet.Cells[i, 1].PutValue($"Item {i - 1}");    // Name
            worksheet.Cells[i, 2].PutValue((i - 1) * 10);       // Score
        }

        // Freeze the first two columns (A and B)
        // Freeze at column C (index 2) with 0 frozen rows and 2 frozen columns
        worksheet.FreezePanes(0, 2, 0, 2);
        // Alternative using cell name:
        // worksheet.FreezePanes("C1", 0, 2);

        // Create XLS save options for legacy compatibility
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an XLS file
        workbook.Save("FrozenColumns.xls", saveOptions);
    }
}
