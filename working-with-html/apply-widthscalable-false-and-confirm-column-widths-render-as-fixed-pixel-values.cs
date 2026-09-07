// Title: How to set a worksheet column to a fixed pixel width and turn off width scaling with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to assign a specific pixel width to a column and ensures the WidthScalable property is disabled. | Show how to read back the pixel width of the column after applying SetColumnWidthPixel to verify the fixed size.
// Common Searches: Aspose.Cells C# set column width in pixels without scaling | Retrieve exact pixel width of an Excel column after using SetColumnWidthPixel | Disable automatic column width scaling in Aspose.Cells workbook | Fixed pixel column width example for .NET Excel generation
// Tags: SetColumnWidthPixel fixed width Aspose.Cells | GetColumnWidthPixel verify column size | disable width scaling Aspose.Cells .NET | C# fixed pixel column width Excel | Aspose.Cells column width pixel handling

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, applying SetColumnWidthPixel to make column A exactly 100 pixels wide (which disables scaling), reading the pixel value with GetColumnWidthPixel to confirm the fixed size, and saving the file.
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

            // Set column A width to 20 characters (default unit)
            sheet.Cells.SetColumnWidth(0, 20);

            // Set column A width to a fixed pixel value (disables scaling)
            int fixedPixelWidth = 100; // example pixel width
            sheet.Cells.SetColumnWidthPixel(0, fixedPixelWidth);

            // Retrieve the pixel width after setting it
            int pixelWidth = sheet.Cells.GetColumnWidthPixel(0);

            // Confirm the column width is rendered as a fixed pixel value
            Console.WriteLine($"Column A pixel width (fixed): {pixelWidth}");

            // Save the workbook
            workbook.Save("WidthScalableDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
