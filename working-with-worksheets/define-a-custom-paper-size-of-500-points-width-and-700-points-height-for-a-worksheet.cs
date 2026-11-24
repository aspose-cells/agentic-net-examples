// Title: Set a custom worksheet paper size of 500 × 700 points using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet custom paper size 500 points width 700 points height | C# Aspose.Cells custom page dimensions for Excel worksheet | Using PageSetup.CustomPaperSize method to define 500x700 point paper size in Aspose.Cells | How to configure a custom paper size in an Excel workbook with Aspose.Cells .NET | Programmatically set custom paper size for a worksheet in Aspose.Cells
// Tags: Aspose.Cells PageSetup.CustomPaperSize API | C# worksheet page setup dimensions points | Excel workbook custom paper metrics | Aspose.Cells set worksheet page size programmatically | PageSetup custom size example .NET

using Aspose.Cells;
using System;

// The example creates a new Workbook, accesses the first Worksheet, switches its PageSetup.PaperSize to Custom, applies a 500‑point width by 700‑point height using the CustomPaperSize method, and saves the file as 'CustomPaperSize.xlsx' while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the paper size to custom
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;

            // Define custom dimensions: 500 points width, 700 points height
            // Use the overload that accepts width and height in points
            sheet.PageSetup.CustomPaperSize(500.0, 700.0);

            // Save the workbook to a file
            string outputPath = "CustomPaperSize.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
