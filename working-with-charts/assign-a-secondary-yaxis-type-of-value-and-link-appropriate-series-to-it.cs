// Title: Create a column chart with a secondary value axis and bind a series to it using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells column chart with secondary Y axis | link a data series to the secondary axis using Aspose.Cells chart API | set titles and range for primary and secondary value axes in Aspose.Cells | use HasAxis method to verify secondary value axis in Aspose.Cells C# | configure min, max and major unit for secondary value axis in Aspose.Cells chart
// Tags: add secondary axis Aspose.Cells | bind series to secondary axis C# | customize primary and secondary axes Aspose.Cells | dual y-axis column chart Aspose.Cells | check secondary axis existence HasAxis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SecondaryValueAxisDemo
{
    // The example creates a new workbook, adds a column chart, defines a primary series (Units Sold) and a secondary series (Revenue), links the secondary series to a secondary value axis, verifies the axis with HasAxis, and customizes titles, ranges, and formatting for both primary and secondary axes before saving the workbook as an XLSX file.
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
                // Category column
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                // Primary series values (e.g., Units Sold)
                sheet.Cells["B1"].PutValue("Units Sold");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Secondary series values (e.g., Revenue)
                sheet.Cells["C1"].PutValue("Revenue");
                sheet.Cells["C2"].PutValue(30000);
                sheet.Cells["C3"].PutValue(45000);
                sheet.Cells["C4"].PutValue(60000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the primary series (Units Sold) – this will plot on the primary Y axis by default
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Units Sold";

                // Add the secondary series (Revenue) – this will be linked to the secondary Y axis
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Revenue";

                // Link the second series to the secondary Y axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Verify that the secondary value axis now exists
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
                Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

                // Customize the secondary value axis (optional)
                Axis secondaryValueAxis = chart.SecondValueAxis;
                secondaryValueAxis.Title.Text = "Revenue (USD)";
                secondaryValueAxis.Title.Font.IsBold = true;
                secondaryValueAxis.Title.Font.Size = 12;
                secondaryValueAxis.MinValue = 0;
                secondaryValueAxis.MaxValue = 70000;
                secondaryValueAxis.MajorUnit = 10000;
                secondaryValueAxis.DisplayUnit = DisplayUnitType.None;

                // Customize the primary value axis (optional)
                Axis primaryValueAxis = chart.ValueAxis;
                primaryValueAxis.Title.Text = "Units Sold";
                primaryValueAxis.Title.Font.IsBold = true;
                primaryValueAxis.Title.Font.Size = 12;
                primaryValueAxis.MinValue = 0;
                primaryValueAxis.MaxValue = 200;
                primaryValueAxis.MajorUnit = 50;

                // Set category axis title
                chart.CategoryAxis.Title.Text = "Month";

                // Save the workbook
                string outputPath = "SecondaryValueAxisDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved to " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
