// Title: C# – Set Z‑Axis Minimum 0 and Maximum 100 for a 3‑D Column Chart with Aspose.Cells
// Description: Creates a workbook, adds a 3‑D column chart, and uses reflection to access the optional ZAxis property. The example disables automatic scaling and forces the Z‑axis range to 0‑100 before saving the file.
// Keywords: Aspose.Cells C# ZAxis scaling | 3D column chart Aspose.Cells | set Z axis min max .NET | Aspose.Cells ZAxis property | reflection Aspose.Cells | Excel 3D chart depth scaling | ChartAxis MinValue MaxValue | Aspose.Cells ChartType.Column3D
// Common Searches: Aspose.Cells set Z axis range .NET | How to set Z axis minimum in 3D chart using C# | ZAxis property missing Aspose.Cells version | Configure Z axis scaling for Column3D chart | C# example for Z axis scaling Aspose.Cells
// Developer Intent: Set a fixed Z‑axis scale (0‑100) on a 3‑D column chart using Aspose.Cells for .NET.
// Use Cases: Standardize depth scaling across multiple 3‑D charts for consistent visual comparison. | Apply reflection to safely configure ZAxis when the property is unavailable in older Aspose.Cells releases. | Generate Excel reports where the Z‑axis range is locked to improve readability of 3‑D data visualizations.
// AI Prompts: Provide C# code that sets the Z‑axis minimum to 0 and maximum to 100 for a 3‑D chart in Aspose.Cells without using reflection. | Explain how to detect ZAxis support in different Aspose.Cells versions and fall back to automatic scaling if the property is absent. | Show how to iterate through all charts in a workbook and apply the same Z‑axis scaling using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a 3‑D column chart, and uses reflection to access the optional ZAxis property. The example disables automatic scaling and forces the Z‑axis range to 0‑100 before saving the file.
    class SetZAxisScaling
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a 3‑D chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(50);

                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(20);
                worksheet.Cells["C3"].PutValue(40);
                worksheet.Cells["C4"].PutValue(60);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Attempt to configure Z‑axis scaling using reflection (ZAxis may not exist in older versions)
                try
                {
                    var zAxisProp = chart.GetType().GetProperty("ZAxis");
                    if (zAxisProp != null)
                    {
                        Axis zAxis = zAxisProp.GetValue(chart) as Axis;
                        if (zAxis != null)
                        {
                            zAxis.IsAutomaticMinValue = false;
                            zAxis.IsAutomaticMaxValue = false;
                            zAxis.MinValue = 0;
                            zAxis.MaxValue = 100;
                        }
                        else
                        {
                            Console.WriteLine("ZAxis property exists but returned null.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("ZAxis property is not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error configuring ZAxis: " + ex.Message);
                }

                // Define output file path
                string outputPath = "ZAxisScalingDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
        }
    }
}
