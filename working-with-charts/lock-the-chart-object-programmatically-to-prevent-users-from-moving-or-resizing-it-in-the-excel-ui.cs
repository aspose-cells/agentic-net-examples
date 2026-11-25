// Title: Programmatically lock an Excel chart to prevent moving or resizing using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock chart shape to disable moving in Excel | prevent chart resizing with Aspose.Cells .NET API | set chart object as locked in worksheet using Aspose.Cells | protect worksheet objects to make chart immovable in C# | Aspose.Cells SetLockedProperty ShapeLockType for chart example | lock Excel chart programmatically with Aspose.Cells workbook protection
// Tags: chart shape locking Aspose.Cells | worksheet objects protection Aspose.Cells | SetLockedProperty chart shape Aspose.Cells | chart movement restriction Aspose.Cells | chart resizing restriction Aspose.Cells | Excel chart lock programmatically .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace LockChartExample
{
    // The example creates a workbook, adds a column chart, retrieves its Shape, applies locked properties (Selection and AdjustHandles) via SetLockedProperty, protects the worksheet for objects, and saves the file, resulting in a chart that cannot be moved or resized in the Excel UI.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // Lock the chart so that users cannot move or resize it
                // ------------------------------------------------------------

                // The chart itself is not a Shape, but it is represented as a Shape
                // in the worksheet's Shapes collection. Retrieve that shape and lock it.
                Shape chartShape = sheet.Shapes[chartIndex];
                chartShape.SetLockedProperty(ShapeLockType.Selection, true);      // Prevent selection/movement
                chartShape.SetLockedProperty(ShapeLockType.AdjustHandles, true); // Prevent resizing

                // Protect the worksheet for objects. This makes all locked objects (including the chart) immutable in the UI.
                sheet.Protect(ProtectionType.Objects);

                // Save the workbook
                string outputPath = "LockedChartWorkbook.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'. The chart is locked and cannot be moved or resized.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
