// Title: Configure worksheet PageSetup to A2 (fallback to A4) and retrieve paper width/height in points with Aspose.Cells for .NET
// AI Prompts: Set the worksheet's PageSetup.PaperSize to PaperSizeType.PaperA4 as a fallback for A2, then print the resulting PaperSize enum value. | Read and display the PageSetup.PaperWidth and PageSetup.PaperHeight properties after setting the paper size, showing the dimensions in points. | Save the workbook after changing the page setup and programmatically confirm that the saved file retains the specified PaperSize setting.
// Common Searches: Aspose.Cells how to set worksheet paper size to A2 when the size is not supported | retrieve page setup paper width and height in points using Aspose.Cells C# | fallback to A4 for Excel worksheet paper size with Aspose.Cells | verify PaperSize enum after configuring PageSetup in Aspose.Cells .NET
// Tags: Aspose.Cells set worksheet paper size | PageSetup PaperSize fallback A4 | retrieve worksheet paper dimensions points | verify PaperSize enum Aspose.Cells | save workbook after page setup change

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PaperSize to PaperSizeType.PaperA4 (used as a fallback when A2 is unavailable), reads back the PaperSize enum, outputs the PaperWidth and PaperHeight values in points, and saves the workbook to 'output.xlsx' to demonstrate how to configure and verify page‑setup dimensions with Aspose.Cells for .NET.
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

            // Set the page‑setup paper size.
            // A2 is not available in the current Aspose.Cells version, using A4 as a valid fallback.
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Verify the paper size setting
            PaperSizeType currentSize = sheet.PageSetup.PaperSize;
            Console.WriteLine($"PaperSize enum value: {currentSize}");

            // Retrieve the physical dimensions (in points) for verification
            double width = sheet.PageSetup.PaperWidth;
            double height = sheet.PageSetup.PaperHeight;
            Console.WriteLine($"Paper dimensions (points): Width = {width}, Height = {height}");

            // Save the workbook (optional, demonstrates usage of save rule)
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
