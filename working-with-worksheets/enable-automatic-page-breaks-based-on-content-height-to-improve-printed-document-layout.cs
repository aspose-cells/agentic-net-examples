// Title: Enable automatic page breaks based on content height in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set page setup to fit width and auto break height | How to enable height-driven page breaks in an Excel file using Aspose.Cells .NET | Aspose.Cells example for automatic pagination with FitToPagesWide = 1 | Set print area and let Aspose.Cells create page breaks automatically | Configure Excel worksheet pagination programmatically with Aspose.Cells
// Tags: Aspose.Cells automatic pagination | PageSetup fit width height-driven breaks | C# set print area Aspose.Cells | Excel worksheet pagination .NET | FitToPagesWide configuration Aspose

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, fills the first worksheet with 200 rows of data, configures the PageSetup to fit the content to one page width while allowing unlimited height (FitToPagesTall = 0) so page breaks are generated automatically based on content height, optionally sets a print area, and saves the workbook as an XLSX file.
class AutomaticPageBreaksExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data to demonstrate page breaks
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure page setup: fit to one page width, unlimited height (automatic breaks)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.FitToPagesWide = 1; // Fit content to one page width
            pageSetup.FitToPagesTall = 0; // Allow as many pages tall as needed (height-driven breaks)

            // Optional: set print area if you want to limit the range
            // pageSetup.PrintArea = "A1:J200";

            // Determine output path
            string outputFile = "AutomaticPageBreaks.xlsx";

            // Save the workbook to a file
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
