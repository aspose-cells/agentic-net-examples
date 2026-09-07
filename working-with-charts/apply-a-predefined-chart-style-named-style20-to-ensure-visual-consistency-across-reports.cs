// Title: How to apply the built‑in Style20 to a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# snippet that creates a workbook, adds sample data, inserts a column chart, and sets its Style property to 20 with Aspose.Cells. | Show the steps to format any Aspose.Cells chart with the predefined Style20 index and then save the workbook as an XLSX file. | Provide code that demonstrates assigning a built‑in chart style index to a chart object in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# apply built‑in chart style 20 to column chart example | set predefined chart style index on Aspose.Cells chart .NET | how to use Style20 for charts in Aspose.Cells workbook
// Tags: apply chart style20 Aspose.Cells C# | set chart.Style property Aspose.Cells | column chart built‑in style Aspose.Cells | save workbook with styled chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartStyleDemo
{
    // Creates a new workbook, populates sample data, adds a column chart, applies the built‑in chart style #20 via the Style property, and saves the file as ChartWithStyle20.xlsx.
    public class ApplyChartStyle20
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", false);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply the predefined built‑in chart style #20 (valid values: 1‑48)
            chart.Style = 20;

            // Save the workbook
            workbook.Save("ChartWithStyle20.xlsx");
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ApplyChartStyle20.Run();
                Console.WriteLine("Workbook created successfully: ChartWithStyle20.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
