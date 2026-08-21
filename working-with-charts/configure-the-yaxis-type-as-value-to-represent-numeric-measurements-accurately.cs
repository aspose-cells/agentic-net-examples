// Title: Set Y‑Axis as a Numeric Value Axis in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, adds category and numeric data, inserts a column chart, assigns data ranges, configures the ValueAxis to a linear (non‑logarithmic) scale with no display‑unit scaling, sets a custom Y‑axis title, and saves the file.
// Keywords: Aspose.Cells Y axis value axis C# | Aspose.Cells linear ValueAxis | C# column chart numeric Y axis | set chart Y axis title Aspose.Cells | Aspose.Cells display unit none | configure chart axis Aspose.Cells | Aspose.Cells chart scaling
// Common Searches: Aspose.Cells set Y axis to value axis C# | How to make Y axis linear in Aspose.Cells chart | C# Aspose.Cells column chart Y axis title | Remove logarithmic scaling from chart axis Aspose.Cells | Aspose.Cells ValueAxis display unit none example
// Developer Intent: Configure a chart’s Y‑axis as a numeric (value) axis with a linear scale, no display‑unit scaling, and a custom title using Aspose.Cells in C#.
// Use Cases: Generate a sales report where column heights represent monetary values on a linear Y‑axis. | Create a scientific worksheet that plots measurement data without logarithmic distortion. | Produce a dashboard workbook with a clearly labeled Y‑axis for presentation to stakeholders.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart and set the Y‑axis as a logarithmic value axis. | Show how to change the Y‑axis display unit to millions in an Aspose.Cells bar chart using C#. | Explain how to read, modify, and save the ValueAxis properties of an existing chart in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds category and numeric data, inserts a column chart, assigns data ranges, configures the ValueAxis to a linear (non‑logarithmic) scale with no display‑unit scaling, sets a custom Y‑axis title, and saves the file.
    public class ConfigureYAxisAsValueDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(250);
            worksheet.Cells["B4"].PutValue(380);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Y values (numeric)
            chart.NSeries.CategoryData = "A2:A4";     // X categories

            // Configure the Y‑axis (ValueAxis) to be a numeric axis
            chart.ValueAxis.IsLogarithmic = false;    // Linear scale
            chart.ValueAxis.DisplayUnit = DisplayUnitType.None; // No scaling

            // Set a title for the Y‑axis
            chart.ValueAxis.Title.Text = "Numeric Measurements";

            // Save the workbook to a file
            string outputPath = "ConfigureYAxisAsValueDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
