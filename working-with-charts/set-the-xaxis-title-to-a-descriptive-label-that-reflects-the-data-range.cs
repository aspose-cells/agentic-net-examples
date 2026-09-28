// Title: How to set a descriptive X‑Axis title for a column chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart from a data range and assign a custom X‑axis title with Aspose.Cells in C#. | Make the X‑axis title visible and set its text to reflect the chart's data range (e.g., "Month (Jan‑Mar)") using Aspose.Cells. | Add a chart title, configure both axis titles, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
// Common Searches: aspnet add X axis label to column chart with Aspose.Cells | c# Aspose.Cells make X axis title appear in Excel chart | how to show axis titles in Aspose.Cells generated chart | customize X axis text for Excel column chart using Aspose.Cells C# | Aspose.Cells chart title and axis visibility example .NET
// Tags: Aspose.Cells column chart axis labeling | configure chart axis titles .NET | C# Aspose.Cells Excel chart customization | add chart title and axis labels Aspose.Cells | save workbook with chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills cells with month and sales data, adds a column chart, assigns the data range, sets a visible X‑axis title "Month (Jan‑Mar)", adds a chart title, and saves the file as SetXAxisTitleDemo.xlsx using Aspose.Cells for .NET.
    public class SetXAxisTitleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (e.g., months and sales)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(1800);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories (X‑axis)

                // Set a descriptive title for the X‑axis that reflects the data range
                chart.CategoryAxis.Title.Text = "Month (Jan‑Mar)";
                chart.CategoryAxis.Title.IsVisible = true;

                // (Optional) Set chart title and make it visible
                chart.Title.Text = "Quarterly Sales";
                chart.Title.IsVisible = true;

                // Define output file path
                string outputPath = "SetXAxisTitleDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetXAxisTitleDemo.Run();
        }
    }
}
