// Title: How to Set Y‑Axis Minimum and Maximum Values in an Aspose.Cells Chart (C#)
// Description: This example shows how to create a workbook, add sample data, insert a column chart, access its ValueAxis, turn off automatic scaling, assign MinValue = 5 and MaxValue = 60, and save the file as YAxisMinMaxDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# chart axis scaling | Y axis minimum value | Y axis maximum value | ValueAxis | set chart axis range | disable automatic axis scaling | column chart Aspose.Cells | Excel chart customization
// Common Searches: Aspose.Cells set Y axis min and max | C# Aspose.Cells chart axis scaling example | disable automatic axis values Aspose.Cells | how to fix Y‑axis range in Aspose.Cells chart | Aspose.Cells ValueAxis MinValue MaxValue
// Developer Intent: Define explicit minimum and maximum limits for a chart's Y‑axis.
// Use Cases: Maintain a consistent baseline across multiple reports for easy visual comparison. | Lock the Y‑axis range when data varies but presentation standards require a fixed scale. | Prepare charts for presentations where axis limits must match predefined business thresholds.
// AI Prompts: Generate C# code with Aspose.Cells to set the Y‑axis minimum to 0 and maximum to 100 for a line chart. | Explain how to disable automatic axis scaling and assign custom MinValue and MaxValue for any chart type in Aspose.Cells. | Provide step‑by‑step instructions to retrieve a chart's ValueAxis from an existing workbook and modify its scaling properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to create a workbook, add sample data, insert a column chart, access its ValueAxis, turn off automatic scaling, assign MinValue = 5 and MaxValue = 60, and save the file as YAxisMinMaxDemo.xlsx using Aspose.Cells for .NET.
    class SetYAxisMinMax
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(50);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the Y‑axis (value axis) and set custom scaling
            Axis yAxis = chart.ValueAxis;

            // Disable automatic min/max calculation
            yAxis.IsAutomaticMinValue = false;
            yAxis.IsAutomaticMaxValue = false;

            // Define the desired minimum and maximum values
            yAxis.MinValue = 5;   // Minimum value on Y‑axis
            yAxis.MaxValue = 60;  // Maximum value on Y‑axis

            // Save the workbook to a file
            string outputPath = "YAxisMinMaxDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}
