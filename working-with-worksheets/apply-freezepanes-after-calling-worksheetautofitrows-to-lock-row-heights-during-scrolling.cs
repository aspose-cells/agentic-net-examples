// Title: Freeze the top row after AutoFitRows to keep row height constant while scrolling in Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to auto‑fit all rows in a worksheet, then freeze the first row so its height remains fixed during scrolling, and save the workbook as an .xlsx file. | Show how to call Worksheet.FreezePanes with the four‑parameter overload after Worksheet.AutoFitRows to lock the top row in a .NET Excel workbook.
// Common Searches: Aspose.Cells C# freeze first row after AutoFitRows | keep row height unchanged while scrolling Aspose.Cells workbook | how to apply FreezePanes after AutoFitRows in .NET | preserve top row height during scroll using Aspose.Cells | C# Aspose.Cells freeze panes after auto fitting rows
// Tags: auto-fit rows then freeze panes Aspose.Cells | freeze top row after auto-fit rows C# | preserve row height during scroll Aspose.Cells | Worksheet.FreezePanes usage after AutoFitRows | lock row height in .NET Excel workbook

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds sample data, calls AutoFitRows to adjust row heights, then freezes the first row with FreezePanes(1,0,1,0) so its height stays constant while scrolling, and finally saves the file as Output.xlsx.
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

            // Populate some sample data (optional)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");
            // Add more rows as needed...

            // AutoFit all rows to adjust their heights based on content
            sheet.AutoFitRows();

            // Freeze the first row (row index is zero‑based). 
            // The 4‑parameter overload specifies rows/columns to freeze and the visible area.
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
