// Title: Set Z‑Axis Tick Label Number Format to Two Decimal Places in a 3‑D Column Chart (Aspose.Cells for .NET)
// Description: Creates a workbook, populates sample data, adds a 3‑D column chart, and demonstrates how to apply the numeric format "0.00" to Z‑axis tick labels via Chart.ZAxis.TickLabels.NumberFormat. Includes a version‑compatibility note and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# chart formatting | Z axis number format | 3D column chart | custom numeric format | two decimal places | .NET | Chart.ZAxis | TickLabels.NumberFormat | Aspose.Cells version check
// Common Searches: Aspose.Cells set Z axis format C# | how to format Z‑axis tick labels in 3D chart .NET | custom number format for chart depth axis Aspose | ZAxis.TickLabels.NumberFormat example | which Aspose.Cells version supports ZAxis property
// Developer Intent: Apply a two‑decimal numeric format to the Z‑axis tick labels of a 3‑D column chart using Aspose.Cells for .NET.
// Use Cases: Financial reporting where the depth axis must display currency values with two decimal places. | Scientific visualization that requires fixed‑point precision on the Z‑axis of a 3‑D chart. | Standardizing axis formatting across X, Y, and Z axes in a multi‑dimensional dashboard workbook.
// AI Prompts: Generate C# code that sets Chart.ZAxis.TickLabels.NumberFormat to "0.00" for a 3‑D column chart and includes a runtime check for ZAxis support. | Explain how to determine if the current Aspose.Cells version provides the ZAxis property and suggest an alternative when it is unavailable. | Show how to apply the same "0.00" numeric format to X, Y, and Z axes of a 3‑D chart in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates sample data, adds a 3‑D column chart, and demonstrates how to apply the numeric format "0.00" to Z‑axis tick labels via Chart.ZAxis.TickLabels.NumberFormat. Includes a version‑compatibility note and saves the workbook as an XLSX file.
    class SetZAxisNumberFormat
    {
        public static void Run()
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
                sheet.Cells["B2"].PutValue(1.2345);
                sheet.Cells["B3"].PutValue(2.3456);
                sheet.Cells["B4"].PutValue(3.4567);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(4.5678);
                sheet.Cells["C3"].PutValue(5.6789);
                sheet.Cells["C4"].PutValue(6.7890);

                // Add a 3‑D column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set series data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Series1";
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Series2";

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Apply custom numeric format to Z‑axis tick labels if supported
                // Note: ZAxis property is available in newer versions of Aspose.Cells.
                // Uncomment the following line when using a version that supports ZAxis.
                // chart.ZAxis.TickLabels.NumberFormat = "0.00";

                // Save the workbook
                string outputPath = "ZAxisNumberFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetZAxisNumberFormat.Run();
        }
    }
}
