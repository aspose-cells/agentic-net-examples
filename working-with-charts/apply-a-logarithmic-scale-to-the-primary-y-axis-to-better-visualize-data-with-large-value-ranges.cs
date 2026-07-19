// Title: Apply a Logarithmic Scale to the Primary Y‑Axis of a Column Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts categorical data that spans several orders of magnitude, adds a column chart, and configures the primary Y (value) axis to use a base‑10 logarithmic scale with explicit minimum and maximum limits and a title, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | logarithmic axis | primary Y axis | chart log scale | column chart | set axis min max | log base 10 | Excel chart programming | Aspose.Cells example
// Common Searches: Aspose.Cells set logarithmic Y axis C# | How to enable log scale on chart axis using Aspose.Cells .NET | C# code for column chart with log‑scaled Y axis in Aspose.Cells | Configure axis min and max for logarithmic chart Aspose.Cells | Add title to logarithmic axis Aspose.Cells C#
// Developer Intent: Configure a chart’s primary Y‑axis to display data on a logarithmic scale in a .NET workbook.
// Use Cases: Financial dashboards where values range from tens to millions and a log‑scaled Y axis improves readability. | Scientific reports visualizing exponential growth or decay across multiple orders of magnitude. | Business intelligence sheets that automatically compress wide‑range data into a single column chart.
// AI Prompts: Generate C# code that creates a column chart with a base‑10 logarithmic primary Y axis, including min/max limits, using Aspose.Cells. | Show how to toggle logarithmic scaling on a chart axis at runtime in Aspose.Cells for .NET. | Explain how to add a descriptive title to a logarithmic Y axis and save the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLogarithmicAxisDemo
{
    // This example creates a workbook, inserts categorical data that spans several orders of magnitude, adds a column chart, and configures the primary Y (value) axis to use a base‑10 logarithmic scale with explicit minimum and maximum limits and a title, then saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a wide value range
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["A5"].PutValue("D");

            // Values span several orders of magnitude
            worksheet.Cells["B2"].PutValue(10);          // 10
            worksheet.Cells["B3"].PutValue(100);         // 100
            worksheet.Cells["B4"].PutValue(1000);        // 1,000
            worksheet.Cells["B5"].PutValue(1000000);    // 1,000,000

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Configure the primary Y axis (ValueAxis) to use a logarithmic scale
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsLogarithmic = true;   // Enable logarithmic scaling
            valueAxis.LogBase = 10;           // Set logarithmic base (default is 10)
            valueAxis.MinValue = 1;           // Define minimum value for the axis
            valueAxis.MaxValue = 10000000;    // Define maximum value for the axis

            // Optional: give the axis a title for clarity
            valueAxis.Title.Text = "Logarithmic Value Axis";

            // Save the workbook to a file
            workbook.Save("LogarithmicAxisDemo.xlsx");
        }
    }
}
