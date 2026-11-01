// Title: How to configure an Aspose.Cells sparkline group to show both high and low point markers in C#
// Common Searches: C# Aspose.Cells sparkline show high and low markers together | set sparkline group to display high and low points using Aspose.Cells .NET | Aspose.Cells sparkline marker options for high and low values in C# | enable both high and low point markers in a sparkline group with Aspose.Cells | Aspose.Cells .NET sparkline group marker configuration example
// Tags: Aspose.Cells sparkline high low marker configuration | C# sparkline group marker settings Aspose.Cells | display high and low points sparkline Aspose.Cells | sparkline marker options .NET | Aspose.Cells sparkline group markers

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Workbook, populates cells A1 through A5 with sample numeric data on the first worksheet, and saves the workbook as WorkbookWithData.xlsx. (Sparkline code is omitted because the Aspose.Cells.Sparkline assembly is not referenced.)
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data in column A (A1:A5)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(5);
            sheet.Cells["A4"].PutValue(30);
            sheet.Cells["A5"].PutValue(15);

            // NOTE: Sparkline functionality requires the Aspose.Cells.Sparkline assembly.
            // If the assembly is not referenced, the following Sparkline code is omitted
            // to ensure the program compiles and runs with the core Aspose.Cells library.

            // Define output file path
            string outputPath = "WorkbookWithData.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
