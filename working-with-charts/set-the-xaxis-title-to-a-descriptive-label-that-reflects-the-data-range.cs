// Title: How to set a descriptive X‑axis title for a column chart in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook with month‑sales data, add a column chart, and set a custom category axis title using Aspose.Cells in C#. | Update an existing Aspose.Cells chart to display a visible X‑axis label that reflects the data range, ensuring the title property is enabled. | Generate a .xlsx file containing a column chart where the X‑axis title is defined programmatically and saved with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set category axis title for column chart | how to make X axis label visible in Aspose.Cells generated Excel chart | example of adding custom X axis text to chart with Aspose.Cells .NET | set chart axis title programmatically using Aspose.Cells for .NET | C# Aspose.Cells chart axis title from data range
// Tags: Aspose.Cells set chart X axis title C# | column chart category axis label Aspose.Cells | Excel chart axis title visibility .NET | programmatic axis labeling Aspose.Cells | C# generate chart with custom X axis text

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a workbook, fills month and sales data, adds a column chart, assigns a visible X‑axis (category) title "Months (Jan - May)", and saves the file as ChartWithXAxisTitle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 1200, 1500, 1800, 1300, 1700 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A: Month
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B: Sales
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Set the X‑axis (category axis) title to describe the data range
            chart.CategoryAxis.Title.Text = "Months (Jan - May)";
            chart.CategoryAxis.Title.IsVisible = true; // Ensure the title is displayed

            // Save the workbook with the chart
            string outputPath = "ChartWithXAxisTitle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
