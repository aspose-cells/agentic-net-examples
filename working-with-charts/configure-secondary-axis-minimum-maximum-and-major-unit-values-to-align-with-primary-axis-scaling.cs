// Title: Synchronize secondary axis minimum, maximum, and major unit with primary axis in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# set secondary value axis minimum and maximum | match secondary axis scaling to primary axis in Aspose.Cells chart | copy primary axis major unit to secondary axis using Aspose.Cells | configure secondary axis range for column chart in C# Aspose.Cells | plot series on second axis and align axis limits Aspose.Cells
// Tags: secondary axis range synchronization Aspose.Cells | custom axis min max major unit C# | column chart secondary value axis configuration | plot series on second axis Aspose.Cells | primary axis scaling copy Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SecondaryAxisScalingDemo
{
    // Creates a workbook with a column chart containing primary and secondary series, sets explicit min, max, and major unit values on the primary value axis, then copies those settings to the secondary axis to keep both axes aligned, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                // Categories
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");
                worksheet.Cells["A5"].PutValue("Apr");

                // Primary series values
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(120);
                worksheet.Cells["B4"].PutValue(80);
                worksheet.Cells["B5"].PutValue(150);

                // Secondary series values (different magnitude)
                worksheet.Cells["C2"].PutValue(500);
                worksheet.Cells["C3"].PutValue(1200);
                worksheet.Cells["C4"].PutValue(800);
                worksheet.Cells["C5"].PutValue(1500);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series to the chart
                chart.NSeries.Add("B2:B5", true); // Primary series
                chart.NSeries[0].Name = "Primary Series";

                chart.NSeries.Add("C2:C5", true); // Secondary series
                chart.NSeries[1].Name = "Secondary Series";

                // Set the category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Plot the second series on the secondary value axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // ------------------------------
                // Configure Primary Value Axis
                // ------------------------------
                Axis primaryAxis = chart.ValueAxis;
                primaryAxis.IsAutomaticMinValue = false;
                primaryAxis.MinValue = 0;               // Minimum value for primary axis
                primaryAxis.IsAutomaticMaxValue = false;
                primaryAxis.MaxValue = 200;             // Maximum value for primary axis
                primaryAxis.IsAutomaticMajorUnit = false;
                primaryAxis.MajorUnit = 50;             // Major unit for primary axis

                // Optional: give the primary axis a title
                primaryAxis.Title.Text = "Primary Axis (Units)";

                // ------------------------------
                // Configure Secondary Value Axis
                // ------------------------------
                Axis secondaryAxis = chart.SecondValueAxis;
                // Align secondary axis scaling with primary axis
                secondaryAxis.IsAutomaticMinValue = false;
                secondaryAxis.MinValue = primaryAxis.MinValue;

                secondaryAxis.IsAutomaticMaxValue = false;
                secondaryAxis.MaxValue = primaryAxis.MaxValue;

                secondaryAxis.IsAutomaticMajorUnit = false;
                secondaryAxis.MajorUnit = primaryAxis.MajorUnit;

                // Optional: give the secondary axis a title
                secondaryAxis.Title.Text = "Secondary Axis (Units)";

                // Save the workbook
                string outputPath = "SecondaryAxisScaling.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
