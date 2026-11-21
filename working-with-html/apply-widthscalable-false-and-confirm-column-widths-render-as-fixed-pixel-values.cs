// Title: Set a column width to a fixed pixel value and verify WidthScalable is disabled using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set column width in pixels and check scaling | how to disable column width scaling in an Excel file with Aspose.Cells | retrieve fixed pixel width of a column after SetColumnWidthPixel in .NET | confirm WidthScalable false after setting column width using Aspose.Cells API | C# code to set and validate fixed column pixel width in XLSX with Aspose.Cells
// Tags: SetColumnWidthPixel API Aspose.Cells | disable column width scaling Aspose.Cells | retrieve column pixel width C# | fixed column width Excel Aspose.Cells | WidthScalable property check

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, sets column A's width to 120 pixels with SetColumnWidthPixel (which automatically sets WidthScalable to false), reads the pixel width back using GetColumnWidthPixel to confirm the fixed size, prints the value, and saves the file as FixedPixelWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set column A (index 0) width to a fixed pixel value.
        // Using SetColumnWidthPixel disables width scaling (WidthScalable = false).
        sheet.Cells.SetColumnWidthPixel(0, 120); // 120 pixels

        // Retrieve the pixel width to confirm it is fixed.
        int pixelWidth = sheet.Cells.GetColumnWidthPixel(0);
        Console.WriteLine($"Column A width (fixed pixels): {pixelWidth}");

        // Save the workbook (lifecycle rule: save)
        workbook.Save("FixedPixelWidth.xlsx");
    }
}
