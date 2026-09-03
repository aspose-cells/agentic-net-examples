// Title: Freeze the top row of a worksheet and save the workbook as an XLSB file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, freezes the first worksheet row, and saves it in XLSB format with Aspose.Cells. | Provide a .NET snippet that applies FreezePanes to lock the header row and exports the file as a binary XLSB workbook. | Write a try‑catch example in C# that builds a worksheet, freezes row 1, and writes the result to an XLSB file using Aspose.Cells.
// Common Searches: aspnet freeze first row of Excel sheet and export as xlsb using Aspose.Cells | how to use FreezePanes method to lock header row in Aspose.Cells C# | save workbook with frozen panes in binary XLSB format with Aspose.Cells for .NET
// Tags: freeze panes aspose cells c# | save workbook as xlsb aspose cells | freeze top row xlsb export | binary excel compression aspose | aspnet worksheet freeze header row

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds sample data, freezes the top row using FreezePanes(1,0,1,0), and saves the file as a binary XLSB workbook via SaveFormat.Xlsb, with exception handling.
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

            // Add some sample data (optional)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["A3"].PutValue("Data 2");

            // Freeze the first row (row index 0) so it stays visible while scrolling
            // row = 1 (first row after the frozen area), column = 0 (no column freeze)
            // totalRows = 1 (freeze one row), totalColumns = 0 (no column freeze)
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook in XLSB format for binary compression
            workbook.Save("FrozenRowsWorkbook.xlsb", SaveFormat.Xlsb);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
