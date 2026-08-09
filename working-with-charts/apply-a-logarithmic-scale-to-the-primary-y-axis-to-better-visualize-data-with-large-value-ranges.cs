// Title: Create a Column Chart with Logarithmic Y‑Axis Using Aspose.Cells for .NET (C#)
// Description: This C# example builds a workbook, inserts categories and values from 10 to 10,000,000, adds a column chart, and configures the primary Y (value) axis to use a base‑10 logarithmic scale (min = 1, max = 100,000,000) with a title before saving as an XLSX file.
// Keywords: Aspose.Cells | C# chart logarithmic axis | logarithmic Y axis .NET | column chart log scale | set chart axis base 10 | chart axis min max Aspose | data visualization large range | .NET spreadsheet chart example
// Common Searches: Aspose.Cells enable logarithmic Y axis C# | how to set log base for chart axis in Aspose.Cells | column chart with log scale using Aspose.Cells .NET | set minimum and maximum values for logarithmic axis Aspose | add title to logarithmic axis Aspose.Cells chart
// Developer Intent: Apply a base‑10 logarithmic scale to the primary Y axis of a column chart in Aspose.Cells for .NET to handle data spanning multiple orders of magnitude.
// Use Cases: Financial dashboards showing revenue from thousands to billions on a single chart. | Scientific plots of exponential growth, such as bacterial counts or particle decay. | Engineering reports comparing low‑level sensor readings with high‑level stress values without visual distortion.
// AI Prompts: Generate C# code with Aspose.Cells that creates a line chart and sets a logarithmic Y axis with custom tick intervals. | Show how to configure a secondary Y axis as logarithmic in an Aspose.Cells chart for .NET. | Explain how to programmatically detect when a data series requires a logarithmic axis and apply the appropriate settings using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLogarithmicAxisDemo
{
    // This C# example builds a workbook, inserts categories and values from 10 to 10,000,000, adds a column chart, and configures the primary Y (value) axis to use a base‑10 logarithmic scale (min = 1, max = 100,000,000) with a title before saving as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a wide range of values
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Low");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Medium");
            worksheet.Cells["B3"].PutValue(1000);
            worksheet.Cells["A4"].PutValue("High");
            worksheet.Cells["B4"].PutValue(100000);
            worksheet.Cells["A5"].PutValue("Very High");
            worksheet.Cells["B5"].PutValue(10000000);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Configure the primary Y axis (ValueAxis) to use a logarithmic scale
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsLogarithmic = true;      // Enable logarithmic scaling
            valueAxis.LogBase = 10;              // Set the logarithmic base (default is 10)
            valueAxis.MinValue = 1;              // Define minimum value for the axis
            valueAxis.MaxValue = 100000000;      // Define maximum value for the axis

            // Optional: give the axis a title for clarity
            valueAxis.Title.Text = "Logarithmic Value Axis";

            // Save the workbook to an XLSX file
            workbook.Save("LogarithmicAxisDemo.xlsx");
        }
    }
}
