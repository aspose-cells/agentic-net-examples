// Title: Set the Y‑Axis to a numeric Value axis and turn off logarithmic scaling in an Aspose.Cells column chart (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets the ValueAxis.IsLogarithmic property to false and DisplayUnit to None. | Demonstrate how to programmatically configure a chart's Y‑axis as a standard numeric axis without logarithmic scaling in a .NET workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart Y axis to value type and disable logarithmic scaling | how to make Y axis numeric in an Aspose.Cells column chart .NET | remove display unit from value axis in Aspose.Cells workbook C# | configure value axis properties for a chart using Aspose.Cells | Aspose.Cells chart axis settings for numeric measurements
// Tags: Aspose.Cells chart value axis configuration | C# set chart Y axis numeric | disable logarithmic scaling Aspose.Cells | column chart display unit none Aspose.Cells | Aspose.Cells workbook chart axis settings

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Added for Chart, ChartType, DisplayUnitType

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, inserts a column chart, and configures the chart's Y‑axis (value axis) to be a non‑logarithmic numeric axis with no display unit before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(250);
                worksheet.Cells["B4"].PutValue(370);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the Y‑axis (value axis) to be a numeric (value) axis
                // Ensure it is not logarithmic and use the default display unit (None)
                chart.ValueAxis.IsLogarithmic = false;
                chart.ValueAxis.DisplayUnit = DisplayUnitType.None;

                // Save the workbook
                string outputPath = "YAxisValueTypeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
