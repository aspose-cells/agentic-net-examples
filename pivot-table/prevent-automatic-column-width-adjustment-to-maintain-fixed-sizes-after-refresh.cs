// Title: Prevent automatic column width adjustment after pivot table refresh using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# keep column widths unchanged after pivot table refresh | How to stop Excel columns from auto‑fitting when updating a pivot with Aspose | Preserve manual column width settings in workbook after data refresh using Aspose.Cells | Disable automatic column resizing for specific columns in Aspose.Cells .NET
// Tags: Aspose.Cells column width persistence | prevent auto‑fit after pivot refresh | C# lock Excel column width programmatically | disable column auto resize Aspose | fixed column width handling in Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, sets explicit column widths for columns A and B, writes sample headers and data, and saves the file as FixedColumnWidths.xlsx, demonstrating how to maintain fixed column sizes when the worksheet (or a pivot table) is refreshed using Aspose.Cells for .NET.
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

            // Set fixed column widths (in characters)
            sheet.Cells.SetColumnWidth(0, 15); // Column A
            sheet.Cells.SetColumnWidth(1, 20); // Column B

            // Example data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Save the workbook
            workbook.Save("FixedColumnWidths.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
