// Title: Set a 2‑point dark gray border on a chart area in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook, add sample data, create a column chart, and apply a 2 pt dark gray border to the chart area using Aspose.Cells in C#. | Write C# code that populates cells, inserts a chart, and configures the chart area border weight to 2 points and color to DarkGray with Aspose.Cells. | Produce a .xlsx file where the column chart’s chart area has a 2‑point thick dark gray border, leveraging the Aspose.Cells ChartArea.Border API.
// Common Searches: Aspose.Cells C# set chart area border thickness to 2 points | how to change chart border color to dark gray using Aspose.Cells | example of customizing chart area border weight in a .NET Excel file | C# code for applying a 2 pt border to an Excel chart with Aspose.Cells | set column chart border style Aspose.Cells .NET
// Tags: chart area border thickness Aspose.Cells | chart area border color DarkGray C# | column chart border formatting .NET | Aspose.Cells chart styling example | Excel chart border customization C# | set chart area border weight Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // // Demonstrates creating a workbook, adding sample data, inserting a column chart, and configuring the chart area's border to 2 points thickness with a dark gray color before saving the file.
    public class ChartBorderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure chart area border
                ChartArea chartArea = chart.ChartArea;
                chartArea.Border.WeightPt = 2.0;          // Thickness in points
                chartArea.Border.Color = Color.DarkGray; // Border color

                // Save the workbook
                workbook.Save("ChartBorderDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ChartBorderDemo.Run();
        }
    }
}
