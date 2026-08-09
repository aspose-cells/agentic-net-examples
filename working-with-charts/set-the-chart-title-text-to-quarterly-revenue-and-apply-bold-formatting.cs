// Title: Set Chart Title to "Quarterly Revenue" and Apply Bold Formatting with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, fill quarter‑revenue data, add a column chart, bind the data range, set the chart title text to "Quarterly Revenue", make the title visible, apply bold styling (and optional dark‑blue color), and save the file as QuarterlyRevenueChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart title | C# set chart title | Aspose.Cells bold title | column chart title formatting | Aspose.Cells .NET example | Excel chart title color | programmatic chart title | Aspose.Cells workbook chart | chart title visibility
// Common Searches: Aspose.Cells set chart title C# | How to make chart title bold in Aspose.Cells | Change Excel chart title color with Aspose.Cells | Add column chart and title using Aspose.Cells .NET | Aspose.Cells chart title formatting example
// Developer Intent: Set a chart title to "Quarterly Revenue" and make it bold (optionally colored) in a C# project using Aspose.Cells.
// Use Cases: Generate a financial report Excel file with a column chart that highlights quarterly revenue via a bold, colored title. | Automate creation of presentation‑ready charts in a .NET application, ensuring the title is programmatically styled for emphasis. | Build a reusable library that adds charts to workbooks with consistent title visibility and formatting across multiple datasets.
// AI Prompts: Write C# code with Aspose.Cells to add a column chart, set its title to "Quarterly Revenue", make the title bold, and change the font color to dark blue. | Provide an Aspose.Cells for .NET example that configures chart title visibility, text, font weight, and color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, fill quarter‑revenue data, add a column chart, bind the data range, set the chart title text to "Quarterly Revenue", make the title visible, apply bold styling (and optional dark‑blue color), and save the file as QuarterlyRevenueChart.xlsx using Aspose.Cells for .NET.
    public class SetChartTitleBoldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["B1"].PutValue("Revenue");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["B2"].PutValue(150000);
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["B3"].PutValue(200000);
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B4"].PutValue(180000);
                sheet.Cells["A5"].PutValue("Q4");
                sheet.Cells["B5"].PutValue(220000);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Set the chart title text and make it bold
                chart.Title.Text = "Quarterly Revenue";
                chart.Title.IsVisible = true;             // Ensure the title is displayed
                chart.Title.Font.IsBold = true;           // Apply bold formatting
                chart.Title.Font.Color = Color.DarkBlue; // Optional title color

                // Save the workbook to a file
                workbook.Save("QuarterlyRevenueChart.xlsx");
                Console.WriteLine("Workbook saved as QuarterlyRevenueChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetChartTitleBoldDemo.Run();
        }
    }
}
