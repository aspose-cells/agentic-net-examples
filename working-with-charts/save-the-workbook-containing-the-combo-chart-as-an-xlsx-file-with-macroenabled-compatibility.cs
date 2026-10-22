// Title: Create a column‑line combo chart and save it as a macro‑enabled XLSX file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that builds a combo chart (column series plus a line series) from sample data and saves the workbook as an .xlsm file. | Demonstrate how to add a secondary line series to a column chart and export the workbook with macro‑enabled compatibility using Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# save workbook as .xlsm macro enabled | how to create a column and line combo chart with Aspose.Cells | add secondary line series to a column chart Aspose.Cells example | export chart workbook to macro enabled Excel file using Aspose.Cells .NET
// Tags: create combo chart Aspose.Cells C# | save workbook as xlsm Aspose.Cells | add secondary line series Aspose.Cells chart | set chart titles Aspose.Cells | macro enabled Excel export .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for ChartType enum

namespace AsposeCellsComboChartExample
{
    // The example creates a new workbook, populates it with sample data, adds a column‑line combo chart with titles, and saves the file as a macro‑enabled XLSX (.xlsm) using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a Combo chart (Column + Line) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];

                // Primary series (Column)
                int series1Idx = chart.NSeries.Add("B2:B4", true);
                var series1 = chart.NSeries[series1Idx];
                series1.Name = "Series 1";

                // Secondary series (Line)
                int series2Idx = chart.NSeries.Add("C2:C4", true);
                var series2 = chart.NSeries[series2Idx];
                series2.Name = "Series 2";
                series2.Type = ChartType.Line; // Set series type to Line
                // Note: Setting series on secondary axis is omitted for compatibility with older Aspose.Cells versions

                // Set chart titles
                chart.Title.Text = "Combo Chart Example";
                chart.CategoryAxis.Title.Text = "Month";
                chart.ValueAxis.Title.Text = "Primary Value";

                // Save the workbook as a macro‑enabled XLSX file
                string outputPath = "ComboChart.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
