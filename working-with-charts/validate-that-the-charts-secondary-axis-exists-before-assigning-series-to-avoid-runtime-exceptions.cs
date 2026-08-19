// Title: Validate Secondary Axis Presence Before Plotting Series – Aspose.Cells C# Sample
// Description: Creates a workbook with a column chart, adds primary and secondary data series, checks if a secondary value axis is available using chart.HasAxis, reveals it when missing, assigns the second series safely, and customizes axis titles and scaling before saving the file.
// Keywords: Aspose.Cells | C# chart example | secondary value axis | HasAxis method | PlotOnSecondAxis | column chart validation | Excel runtime safety | axis visibility
// Common Searches: Aspose.Cells check for secondary axis C# | how to plot series on second axis without error | make hidden secondary axis visible Aspose.Cells | chart.HasAxis usage example | prevent exception when assigning secondary series
// Developer Intent: Confirm that a secondary value axis exists and is visible before linking a data series to it, avoiding runtime failures.
// Use Cases: Generate a column chart with two data sets where the second set uses a separate scale. | Automatically reveal a hidden secondary axis for chart types that hide it by default. | Apply conditional logic to handle axis presence across various chart formats.
// AI Prompts: Write C# code with Aspose.Cells that builds a line chart, adds two series, and verifies the secondary axis before assigning the second series. | Show how to enable a hidden secondary value axis in a bar chart and set its title and range using Aspose.Cells. | Explain the purpose of chart.HasAxis and how it prevents exceptions when using PlotOnSecondAxis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisCheck
{
    // Creates a workbook with a column chart, adds primary and secondary data series, checks if a secondary value axis is available using chart.HasAxis, reveals it when missing, assigns the second series safely, and customizes axis titles and scaling before saving the file.
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

            // Add primary series data
            chart.NSeries.Add("B2:B4", true);
            // Add secondary series data
            chart.NSeries.Add("C2:C4", true);
            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Validate that a secondary value axis exists before assigning the series to it
            bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
            if (hasSecondaryValueAxis)
            {
                // Plot the second series on the secondary axis
                chart.NSeries[1].PlotOnSecondAxis = true;
            }
            else
            {
                // If the secondary axis does not exist, make it visible (some chart types hide it by default)
                chart.SecondValueAxis.IsVisible = true;
                // Now it is safe to assign the series to the secondary axis
                chart.NSeries[1].PlotOnSecondAxis = true;
            }

            // Optional: customize the secondary axis appearance
            Axis secAxis = chart.SecondValueAxis;
            secAxis.Title.Text = "Secondary Values";
            secAxis.MinValue = 0;
            secAxis.MaxValue = 600;
            secAxis.MajorUnit = 100;

            // Save the workbook
            workbook.Save("ChartWithValidatedSecondaryAxis.xlsx");
        }
    }
}
