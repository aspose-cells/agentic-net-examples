// Title: Set Y‑Axis to a Linear Value Axis in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, adds category and numeric data, inserts a column chart, and configures the chart’s ValueAxis as a linear numeric axis by disabling logarithmic scaling, clearing display‑unit conversion, and assigning a custom title before saving.
// Keywords: Aspose.Cells Y axis | C# chart numeric axis | linear value axis Aspose.Cells | disable logarithmic scaling chart | display unit none Aspose.Cells | column chart axis settings | Excel chart value axis C# | .NET chart axis configuration | chart title Aspose.Cells | Excel export numeric measurements
// Common Searches: Aspose.Cells set Y axis to numeric C# | how to make chart Y axis linear in Aspose.Cells | disable logarithmic scaling for chart axis .NET | remove display unit from Y axis Aspose.Cells | add title to chart value axis C#
// Developer Intent: Configure a chart’s Y‑axis as a numeric (value) axis to display raw measurements without logarithmic scaling or unit conversion.
// Use Cases: Generate a sales column chart where the Y‑axis shows exact figures on a linear scale. | Produce an engineering report in Excel that requires a plain numeric Y‑axis for sensor readings. | Create a financial dashboard where the Y‑axis must remain linear and include a descriptive title.
// AI Prompts: Write C# code with Aspose.Cells to set a column chart’s Y‑axis to a linear value axis and add a custom title. | Show how to turn off logarithmic scaling and set DisplayUnit to None for the Y‑axis of a bar chart using Aspose.Cells. | Explain how to programmatically verify that a chart’s Y‑axis type is Value (linear) in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    // Creates a workbook, adds category and numeric data, inserts a column chart, and configures the chart’s ValueAxis as a linear numeric axis by disabling logarithmic scaling, clearing display‑unit conversion, and assigning a custom title before saving.
    class Program
    {
        static void Main(string[] args)
        {
            try
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
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);
                sheet.Cells["B4"].PutValue(370);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the Y‑axis (value axis) to be a numeric (linear) axis
                chart.ValueAxis.IsLogarithmic = false;                 // Ensure linear scaling
                chart.ValueAxis.DisplayUnit = DisplayUnitType.None;    // No display unit conversion
                chart.ValueAxis.Title.Text = "Numeric Measurements";   // Optional title for clarity

                // Save the workbook
                string outputPath = "YAxisValueTypeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
