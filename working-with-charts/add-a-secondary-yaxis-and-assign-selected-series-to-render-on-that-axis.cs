// Title: Create a column chart with a secondary Y‑axis and assign a series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add secondary Y axis to a column chart | Assign a data series to the secondary value axis in Aspose.Cells .NET | Set IsSecondaryValueAxis property via reflection in Aspose.Cells chart | Customize secondary Y‑axis range and title for a column chart using Aspose.Cells | Create dual Y‑axis column chart and export to XLSX with Aspose.Cells for .NET
// Tags: add secondary y axis Aspose.Cells | assign series to secondary value axis .NET | reflection set IsSecondaryValueAxis Aspose.Cells | customize secondary axis properties Aspose.Cells | dual y axis column chart C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AddSecondaryYAxisExample
{
    // The example builds a new workbook, fills it with sample data, and inserts a column chart. It adds a primary series (column B) and a secondary series (column C), then uses reflection to enable the IsSecondaryValueAxis flag, causing Aspose.Cells to generate a secondary Y‑axis. Both axes are customized with titles, fixed min/max values, and font styling. Series names are set, and the workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Primary");
                sheet.Cells["C1"].PutValue("Secondary");

                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(60);
                sheet.Cells["C5"].PutValue(80);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Primary series (Column B)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Secondary series (Column C)
                chart.NSeries.Add("C2:C5", true);
                Series secondarySeries = chart.NSeries[1];

                // Attempt to mark the series for secondary Y‑axis using reflection
                var secondaryProp = secondarySeries.GetType().GetProperty("IsSecondaryValueAxis");
                if (secondaryProp != null && secondaryProp.CanWrite)
                {
                    secondaryProp.SetValue(secondarySeries, true);
                }

                // Verify that the secondary axis now exists
                bool hasSecondary = chart.HasAxis(AxisType.Value, false);
                Console.WriteLine("Secondary Y‑axis created: " + hasSecondary);

                // Customize the primary Y‑axis
                Axis primaryAxis = chart.ValueAxis;
                primaryAxis.Title.Text = "Primary Values";
                primaryAxis.IsAutomaticMaxValue = false;
                primaryAxis.MaxValue = 250;
                primaryAxis.IsAutomaticMinValue = false;
                primaryAxis.MinValue = 0;

                // Customize the secondary Y‑axis if it exists
                if (hasSecondary)
                {
                    var secAxisProp = chart.GetType().GetProperty("SecondaryValueAxis");
                    if (secAxisProp != null)
                    {
                        Axis secondaryAxis = secAxisProp.GetValue(chart) as Axis;
                        if (secondaryAxis != null)
                        {
                            secondaryAxis.Title.Text = "Secondary Values";
                            secondaryAxis.IsAutomaticMaxValue = false;
                            secondaryAxis.MaxValue = 100;
                            secondaryAxis.IsAutomaticMinValue = false;
                            secondaryAxis.MinValue = 0;
                            secondaryAxis.Title.Font.Size = 10;
                            secondaryAxis.Title.Font.Color = Color.DarkGreen;
                        }
                    }
                }

                // Optional: give each series a distinct name
                chart.NSeries[0].Name = "Primary Series";
                chart.NSeries[1].Name = "Secondary Series";

                // Save the workbook
                string outFile = "ChartWithSecondaryYAxis.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
