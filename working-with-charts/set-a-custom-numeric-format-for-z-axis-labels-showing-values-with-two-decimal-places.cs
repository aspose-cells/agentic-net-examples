// Title: C# – Set Z‑Axis Tick Labels to Two Decimal Places in a 3‑D Chart with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a 3‑D column chart, and uses reflection to access the ZAxis (when available). The code sets ZAxis.TickLabels.NumberFormat to "0.00" so Z‑axis values appear with two decimal places, then saves the file.
// Keywords: Aspose.Cells ZAxis format | C# 3D chart numeric format | set Z axis tick label format | reflection Aspose.Cells ZAxis | custom number format chart axis | two decimal places Excel chart | Aspose.Cells .NET chart axis formatting
// Common Searches: how to format Z axis labels in Aspose.Cells | set numeric format for ZAxis tick labels C# | Aspose.Cells 3D chart ZAxis reflection example | display two decimal places on chart axis Aspose | C# chart axis number format Aspose.Cells
// Developer Intent: Apply a custom numeric format (two decimal places) to the Z‑axis tick labels of a 3‑D chart using Aspose.Cells for .NET, with fallback via reflection for older library versions.
// Use Cases: Generate a 3‑D column chart where the depth (Z) values need precise two‑decimal display. | Maintain compatibility across different Aspose.Cells releases by using reflection to access ZAxis only when it exists. | Extend the same reflection technique to modify XAxis or YAxis formatting when direct properties are unavailable.
// AI Prompts: Write C# code that sets ZAxis.TickLabels.NumberFormat to "0.00" in an Aspose.Cells 3‑D chart, using reflection to avoid compile‑time dependency. | Explain how to detect the ZAxis property at runtime and safely apply a numeric format for older Aspose.Cells versions. | Provide a pattern for applying custom number formats to XAxis, YAxis, and ZAxis tick labels in Aspose.Cells charts via reflection.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a 3‑D column chart, and uses reflection to access the ZAxis (when available). The code sets ZAxis.TickLabels.NumberFormat to "0.00" so Z‑axis values appear with two decimal places, then saves the file.
    public class ZAxisTickLabelsNumberFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for a 3‑D chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["B2"].PutValue(1.2345);
            worksheet.Cells["B3"].PutValue(2.3456);
            worksheet.Cells["B4"].PutValue(3.4567);

            // Add a 3‑D column chart (Z axis will be present)
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Attempt to set custom numeric format for Z‑axis tick labels (two decimal places)
            // Use reflection to avoid compile‑time dependency on ZAxis property (may not exist in older versions)
            PropertyInfo zAxisProp = chart.GetType().GetProperty("ZAxis");
            if (zAxisProp != null)
            {
                Axis zAxis = zAxisProp.GetValue(chart) as Axis;
                if (zAxis != null && zAxis.TickLabels != null)
                {
                    zAxis.TickLabels.NumberFormat = "0.00";
                }
            }

            // Define output file path
            string outputPath = "ZAxisTickLabelsNumberFormatDemo.xlsx";

            // Save the workbook (overwrite if it already exists)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
