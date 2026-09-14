// Title: Set a bold, visible title for a column chart in Aspose.Cells using C#
// AI Prompts: Create a new workbook, populate it with quarterly revenue data, add a column chart, set the chart title to "Quarterly Revenue", make the title visible, and apply bold font styling with Aspose.Cells for .NET. | Modify an existing Aspose.Cells chart by changing its title text, enabling title visibility, and configuring the title font to be bold, size 14, and black in C#. | Generate an Excel file that contains a column chart whose title is formatted with custom font properties (bold, specific size, color) using the Aspose.Cells API in C#.
// Common Searches: C# code to make an Aspose.Cells chart title bold and visible | how to change font style of a chart title in Aspose.Cells .NET | set custom title text and formatting for column chart using Aspose.Cells C# | Aspose.Cells example for formatting chart title font properties
// Tags: Aspose.Cells chart title bold formatting | column chart title font customization Aspose.Cells | set chart title visibility Aspose.Cells .NET | apply custom font size color to Excel chart title C# | Aspose.Cells set chart title text programmatically

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding quarterly revenue data, inserting a column chart, setting the chart title to "Quarterly Revenue", making the title visible, and applying bold, size‑14, black font before saving the Excel file.
    public class SetChartTitleBoldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B1"].PutValue("Revenue");
                sheet.Cells["B2"].PutValue(15000);
                sheet.Cells["B3"].PutValue(20000);
                sheet.Cells["B4"].PutValue(18000);
                sheet.Cells["B5"].PutValue(22000);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Set chart title text and make it visible
                chart.Title.Text = "Quarterly Revenue";
                chart.Title.IsVisible = true;

                // Apply bold formatting and other font properties to the title
                chart.Title.Font.IsBold = true;
                chart.Title.Font.Size = 14;
                chart.Title.Font.Color = Color.Black;

                // Save the workbook
                string outputPath = "QuarterlyRevenueChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartTitleBoldDemo.Run();
        }
    }
}
