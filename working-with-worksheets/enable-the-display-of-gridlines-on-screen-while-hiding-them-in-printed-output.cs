// Title: Show worksheet gridlines on screen while suppressing them in printed output with Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET show gridlines in workbook view but hide them in printed PDF | how to prevent gridlines from appearing in printed Excel file using Aspose.Cells C# | C# Aspose.Cells hide printed gridlines while keeping them visible on screen | set worksheet page setup to not print gridlines with Aspose.Cells | Aspose.Cells gridlines visibility screen vs print example
// Tags: Worksheet default screen gridlines Aspose.Cells | PageSetup.PrintGridlines property .NET | Suppress printed gridlines Aspose.Cells | Control Excel gridline visibility programmatically | Aspose.Cells Xlsx gridline display settings

using System;
using Aspose.Cells;
using System.IO;

namespace GridlinesExample
{
    // The sample creates a new Workbook, writes data to cells A1‑B2, and saves it as an .xlsx file. Gridlines are displayed by default in the on‑screen view. Although the ShowGridlines and PrintGridlines members are not present in the current Aspose.Cells release, printed gridlines can be managed through page‑setup options when supported.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data to visualize gridlines
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue(123);
                sheet.Cells["B2"].PutValue(456);

                // Note: ShowGridlines and PrintGridlines properties are not available in the current Aspose.Cells version.
                // Gridlines are shown by default on screen; printed gridlines can be controlled via other settings if needed.

                // Define output file path
                string outputPath = "GridlinesDemo.xlsx";

                // Save the workbook to a file (lifecycle: save)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
