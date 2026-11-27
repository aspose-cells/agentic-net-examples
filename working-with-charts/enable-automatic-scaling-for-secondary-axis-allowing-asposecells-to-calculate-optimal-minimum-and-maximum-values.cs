// Title: How to enable automatic scaling of the secondary value axis in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# set secondary axis automatic min and max values | Enable auto scaling for secondary value axis in an Aspose.Cells chart | C# Aspose.Cells column chart secondary axis automatic scaling example | Make secondary axis visible and let Aspose.Cells calculate limits | Use chart.Calculate to apply automatic scaling on secondary axis in C#
// Tags: secondary value axis auto scaling Aspose.Cells | C# Aspose.Cells column chart secondary axis | chart.Calculate to resolve axis limits Aspose.Cells | make secondary axis visible Aspose.Cells | automatic major and minor units Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAutomaticSecondaryAxisScaling
{
    // The example creates a workbook, adds category, revenue, and units‑sold data, builds a column chart, assigns revenue to the primary axis and units sold to the secondary axis, makes the secondary axis visible, enables automatic minimum, maximum, major and minor units, calls chart.Calculate to compute the optimal axis limits, prints the calculated values, and saves the file as AutomaticSecondaryAxisScaling.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // Populate sample data
                // ------------------------------------------------------------
                // Category (X) values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                // Primary series values (left Y axis)
                sheet.Cells["B1"].PutValue("Revenue");
                sheet.Cells["B2"].PutValue(12000);
                sheet.Cells["B3"].PutValue(15000);
                sheet.Cells["B4"].PutValue(18000);
                sheet.Cells["B5"].PutValue(21000);

                // Secondary series values (right Y axis)
                sheet.Cells["C1"].PutValue("Units Sold");
                sheet.Cells["C2"].PutValue(300);
                sheet.Cells["C3"].PutValue(450);
                sheet.Cells["C4"].PutValue(500);
                sheet.Cells["C5"].PutValue(650);

                // ------------------------------------------------------------
                // Add a column chart
                // ------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the primary series (Revenue)
                chart.NSeries.Add("B2:B5", true);
                // Add the secondary series (Units Sold)
                chart.NSeries.Add("C2:C5", true);
                // Set the category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // ------------------------------------------------------------
                // Enable automatic scaling for the secondary value axis
                // ------------------------------------------------------------
                Axis secondaryValueAxis = chart.SecondValueAxis;

                // Ensure the secondary axis is visible (some chart types hide it by default)
                secondaryValueAxis.IsVisible = true;

                // Let Excel/Aspose.Cells calculate optimal min/max values automatically
                secondaryValueAxis.IsAutomaticMinValue = true;
                secondaryValueAxis.IsAutomaticMaxValue = true;
                secondaryValueAxis.IsAutomaticMajorUnit = true;
                secondaryValueAxis.IsAutomaticMinorUnit = true;

                // ------------------------------------------------------------
                // Force chart calculation so that automatic values are resolved
                // ------------------------------------------------------------
                chart.Calculate();

                // After calculation, the axis objects contain the computed values.
                Console.WriteLine("Secondary Axis Automatic Scaling Results:");
                Console.WriteLine($"  MinValue : {secondaryValueAxis.MinValue}");
                Console.WriteLine($"  MaxValue : {secondaryValueAxis.MaxValue}");
                Console.WriteLine($"  MajorUnit: {secondaryValueAxis.MajorUnit}");
                Console.WriteLine($"  MinorUnit: {secondaryValueAxis.MinorUnit}");

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "AutomaticSecondaryAxisScaling.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors – do not hide failures
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
