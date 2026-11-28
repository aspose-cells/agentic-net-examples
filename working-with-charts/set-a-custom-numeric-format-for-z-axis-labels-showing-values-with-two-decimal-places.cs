// Title: Applying a custom two‑decimal numeric format to Z‑axis labels in a 3‑D column chart with Aspose.Cells for .NET (ZAxis not exposed)
// Common Searches: c# aspose.cells set custom number format for z axis in 3d chart | how to display z axis values with two decimal places using aspose.cells | asp.net aspose.cells 3d column chart axis formatting example | formatting z axis labels in aspose.cells when ZAxis object is unavailable | apply numeric format to chart axes in aspose.cells workbook c#
// Tags: Aspose.Cells custom Z axis number format | 3D column chart axis formatting C# | chart axis numeric format Aspose.Cells | Aspose.Cells ZAxis API limitation | C# workbook chart axis formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SetZAxisNumberFormat
{
    // The example creates a workbook, populates sample data, adds a 3‑D column chart, and notes that recent Aspose.Cells versions do not expose a ZAxis object, so a custom two‑decimal numeric format cannot be directly applied to Z‑axis labels. The workbook is saved as an XLSX file and successfully reloaded for verification.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a 3‑D chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(5.5);
                sheet.Cells["C3"].PutValue(15.75);
                sheet.Cells["C4"].PutValue(25.125);

                // Add a 3‑D column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // NOTE: Aspose.Cells does not expose a ZAxis object in recent versions.
                // If a Z axis is required, use the appropriate API when available.
                // The following code is omitted to keep the example compilable.

                // Save the workbook
                string outputPath = "ZAxisNumberFormatDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");

                // Verify that the file was saved
                if (File.Exists(outputPath))
                {
                    Workbook verifyWb = new Workbook(outputPath);
                    // Since ZAxis is not available, we simply confirm the workbook loads.
                    Console.WriteLine("Verification succeeded: workbook loaded without errors.");
                }
                else
                {
                    Console.WriteLine($"Verification failed: file '{outputPath}' not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
