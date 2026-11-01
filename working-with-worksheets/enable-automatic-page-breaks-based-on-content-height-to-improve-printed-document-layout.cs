// Title: Automatically insert vertical page breaks in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet to auto‑fit vertical pages for printing | How to let Aspose.Cells calculate page breaks based on row height | Generate Excel file with automatic page breaks using Aspose.Cells .NET | Aspose.Cells FitToPagesTall = 0 example for dynamic pagination | Print‑ready Excel workbook with auto vertical page breaks in C#
// Tags: FitToPagesTall property automatic pagination Aspose.Cells | vertical page break generation Excel .NET | auto page layout optimization Aspose.Cells C# | dynamic pagination based on content height | Excel workbook printing setup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new workbook, fills 200 rows with data, sets the worksheet's PageSetup.FitToPagesTall property to 0 so Aspose.Cells automatically determines vertical page breaks according to content height, and saves the result as AutomaticPageBreaks.xlsx.
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

            // Populate sample data to illustrate page breaks
            for (int i = 0; i < 200; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue(i * 10);
            }

            // Let Aspose.Cells determine the number of pages vertically
            sheet.PageSetup.FitToPagesTall = 0; // 0 = automatic

            // Define output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutomaticPageBreaks.xlsx");

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
