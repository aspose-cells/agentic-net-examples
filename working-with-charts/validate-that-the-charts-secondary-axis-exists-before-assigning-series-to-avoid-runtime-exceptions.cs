// Title: Validate Chart Secondary Axis Existence Before Assigning Series in Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a column chart with primary and secondary data series, checks if a secondary value axis is present using `chart.HasAxis(AxisType.Value, false)`, and only sets `PlotOnSecondAxis` for the second series when the axis exists. It also shows optional customization of the secondary axis and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | chart secondary axis | HasAxis | PlotOnSecondAxis | secondary value axis | runtime exception prevention | chart axis validation | Aspose.Cells chart example
// Common Searches: Aspose.Cells check secondary axis before PlotOnSecondAxis | How to verify secondary value axis exists in Aspose.Cells C# | Prevent InvalidOperationException when assigning series to secondary axis Aspose.Cells | C# Aspose.Cells chart secondary axis validation | Add secondary axis to column chart Aspose.Cells
// Developer Intent: Confirm that a chart contains a secondary value axis before assigning a series to it, avoiding runtime errors.
// Use Cases: Safely map a data series to the secondary axis only when the axis is present. | Conditionally format or label the secondary axis after validation. | Log a warning or apply fallback logic when the secondary axis is missing. | Create a reusable helper method that detects a secondary axis in Aspose.Cells charts.
// AI Prompts: Generate C# code that adds a column chart with primary and secondary series using Aspose.Cells, checks `chart.HasAxis(AxisType.Value, false)`, and sets `PlotOnSecondAxis` only if the secondary axis exists. | Write a C# utility method that receives a `Chart` object and returns true if a secondary value axis is available, otherwise logs a warning. | Show an example that first adds a secondary axis to a chart, then safely assigns a series to it with validation, including optional axis title and visibility settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisValidation
{
    // This C# example creates a workbook, adds a column chart with primary and secondary data series, checks if a secondary value axis is present using `chart.HasAxis(AxisType.Value, false)`, and only sets `PlotOnSecondAxis` for the second series when the axis exists. It also shows optional customization of the secondary axis and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Primary Series");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Secondary Series");
            sheet.Cells["C2"].PutValue(500);
            sheet.Cells["C3"].PutValue(300);
            sheet.Cells["C4"].PutValue(100);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set primary series data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add secondary series data
            chart.NSeries.Add("C2:C4", true);

            // Validate that the secondary value axis exists before assigning PlotOnSecondAxis
            bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
            if (hasSecondaryValueAxis)
            {
                // Assign the second series to the secondary axis safely
                chart.NSeries[1].PlotOnSecondAxis = true;
            }
            else
            {
                // If the secondary axis does not exist, you may choose to skip the assignment
                // or handle it according to your business logic. Here we simply output a message.
                Console.WriteLine("Secondary value axis does not exist; PlotOnSecondAxis not set.");
            }

            // Optional: customize secondary axis appearance if it exists
            if (hasSecondaryValueAxis)
            {
                Axis secondaryAxis = chart.SecondValueAxis;
                secondaryAxis.Title.Text = "Secondary Axis";
                secondaryAxis.IsVisible = true;
            }

            // Save the workbook
            workbook.Save("ChartWithValidatedSecondaryAxis.xlsx");
        }
    }
}
