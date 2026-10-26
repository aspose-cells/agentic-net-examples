// Title: Create a custom worksheet view named “ReportView” with frozen header rows and preserved view settings using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to create a named view and freeze the first row | freeze header rows while keeping gridlines and ruler visible in Aspose.Cells workbook | set custom worksheet view properties and save as ReportView.xlsx with Aspose.Cells | C# Aspose.Cells FreezePanes example with view settings preservation | create worksheet view and maintain view options like gridlines in Aspose.Cells for .NET
// Tags: custom worksheet view Aspose.Cells | freeze header rows Aspose.Cells C# | preserve gridlines ruler Aspose.Cells | set view properties Aspose.Cells workbook | save workbook with named view Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to create a new workbook, define a custom view called "ReportView", freeze the header row using FreezePanes, keep gridlines, row/column headers, and ruler visible, and save the file as ReportView.xlsx with Aspose.Cells for .NET.
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

                // Freeze the first row (header row)
                // Parameters: totalRows, totalColumns, row index of first unfrozen row, column index of first unfrozen column
                sheet.FreezePanes(1, 0, 1, 0);

                // Optional view settings
                sheet.IsGridlinesVisible = true;
                sheet.IsRowColumnHeadersVisible = true;
                sheet.IsRulerVisible = true;

                // Define output file path
                string outputFile = "ReportView.xlsx";

                // Ensure the directory exists
                string fullPath = Path.GetFullPath(outputFile);
                string outputDir = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to: {fullPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
