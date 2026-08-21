// Title: Format Y‑Axis as Currency with Two Decimals in an Aspose.Cells Chart (C#)
// Description: This example creates a workbook, adds a column chart, and sets the ValueAxis tick‑label format to "$#,##0.00" so the Y‑axis displays monetary values with two decimal places before saving the file.
// Keywords: Aspose.Cells Y axis currency | C# chart axis number format | Excel chart value axis format | currency tick labels Aspose.Cells | .NET chart formatting | financial chart formatting | Excel number format for axis
// Common Searches: Aspose.Cells set Y axis to currency C# | format chart axis as $ with two decimals | C# Aspose.Cells number format for value axis | how to display monetary values on chart Y axis | apply custom number format to Excel chart axis .NET
// Developer Intent: Apply a currency number format with two decimal places to the Y‑axis (ValueAxis) tick labels of an Aspose.Cells chart using C#.
// Use Cases: Generate financial dashboards where column charts show dollar amounts with precise cent values. | Create sales performance reports that automatically format Y‑axis values as US dollars for stakeholder clarity. | Export budgeting spreadsheets with charts that present monetary figures in a standard currency style.
// AI Prompts: Write C# code with Aspose.Cells to format the Y‑axis of a line chart as Euro currency with three decimal places. | Show how to set custom number formats for both X‑axis and Y‑axis tick labels in an Aspose.Cells chart. | Explain how to read existing axis number formats from an Excel file and update them to a specified currency format using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart, and sets the ValueAxis tick‑label format to "$#,##0.00" so the Y‑axis displays monetary values with two decimal places before saving the file.
    public class YAxisCurrencyFormatDemo
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
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1234.56);
                worksheet.Cells["B3"].PutValue(2345.67);
                worksheet.Cells["B4"].PutValue(3456.78);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Format the Y‑axis (ValueAxis) tick labels as currency with two decimal places
                chart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

                // Save the workbook to a file
                workbook.Save("YAxisCurrencyFormatDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
