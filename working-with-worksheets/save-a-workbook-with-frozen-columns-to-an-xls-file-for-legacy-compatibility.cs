// Title: Freeze the first two columns of a worksheet and save as an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, freeze columns A and B on the first sheet with FreezePanes, then save the file as an Excel97To2003 (.xls) workbook using Aspose.Cells for .NET. | Add sample data, apply FreezePanes to lock the first two columns, and export the workbook to legacy .xls format via the SaveFormat.Excel97To2003 option.
// Common Searches: asp.net freeze first two columns and save as xls using Aspose.Cells | c# Aspose.Cells freeze panes then export to Excel 97-2003 format | how to create frozen columns in an .xls file with Aspose.Cells for .NET | save workbook with frozen columns to legacy Excel format c# Aspose.Cells | Aspose.Cells FreezePanes example for Excel97To2003 output
// Tags: FreezePanes Aspose.Cells C# | export to Excel97To2003 format Aspose.Cells | freeze columns A B Aspose.Cells | legacy .xls workbook generation Aspose.Cells | save workbook with frozen panes C#

using System;
using Aspose.Cells;

// The example creates a new workbook, uses FreezePanes to lock the first two columns on the first worksheet, adds sample header values, and saves the result as an Excel 97‑2003 (.xls) file for legacy compatibility.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first two columns (A and B)
            // Parameters: row, column, totalRows, totalColumns.
            // Row = 0 (no rows frozen), Column = 2 (freeze columns before C).
            sheet.FreezePanes(0, 2, 0, 0);

            // Add sample data to visualize the frozen columns
            sheet.Cells["A1"].PutValue("Frozen Col 1");
            sheet.Cells["B1"].PutValue("Frozen Col 2");
            sheet.Cells["C1"].PutValue("Scrollable Col 3");

            // Save the workbook in legacy XLS format for compatibility
            workbook.Save("output.xls", SaveFormat.Excel97To2003);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
