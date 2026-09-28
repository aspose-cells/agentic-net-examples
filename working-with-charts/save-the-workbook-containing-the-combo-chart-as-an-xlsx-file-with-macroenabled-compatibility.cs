// Title: Generate a column‑line combo chart in a worksheet and save the workbook as a macro‑enabled XLSX (.xlsm) using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to fill a worksheet with sample data, creates a combo chart with a column series and a line series on a secondary axis, and saves the file as an .xlsm workbook. | Show how to assign category labels, set a chart title, and place a line series on a secondary axis in an Aspose.Cells combo chart before exporting to macro‑enabled XLSX. | Demonstrate persisting an Aspose.Cells workbook that contains a chart to the Xlsm format while preserving all chart configurations.
// Common Searches: Aspose.Cells C# create combo chart with column and line series and save as .xlsm | how to add secondary axis to line series in Aspose.Cells chart | saving Aspose.Cells workbook with charts to macro enabled XLSX file
// Tags: combo chart creation Aspose.Cells C# | save workbook as xlsm Aspose.Cells | secondary axis line series Aspose.Cells | populate worksheet data ranges Aspose.Cells | set chart title Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with sample data, adds a combo chart that combines a column series and a line series on a secondary axis, sets chart titles and category labels, and saves the result as a macro‑enabled XLSX (.xlsm) file using Aspose.Cells for .NET.
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

            // Populate header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            // Sample data
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] series1 = { 10, 20, 30, 40, 50 };
            double[] series2 = { 5, 15, 25, 35, 45 };

            // Fill data into cells
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(series1[i]);    // Column B
                sheet.Cells[i + 1, 2].PutValue(series2[i]);    // Column C
            }

            // Add a Combo chart (Column + Line) to the worksheet
            // Position: from row 7, column 0 to row 20, column 7
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart comboChart = sheet.Charts[chartIndex];

            // First series as Column
            comboChart.NSeries.Add("B2:B6", true);
            comboChart.NSeries[0].Name = "Series 1";
            comboChart.NSeries[0].Type = ChartType.Column;

            // Second series as Line (secondary axis)
            comboChart.NSeries.Add("C2:C6", true);
            comboChart.NSeries[1].Name = "Series 2";
            comboChart.NSeries[1].Type = ChartType.Line;
            comboChart.NSeries[1].PlotOnSecondAxis = true; // Use secondary axis for line series

            // Set category axis labels
            comboChart.NSeries.CategoryData = "A2:A6";

            // Optional: give the chart a title
            comboChart.Title.Text = "Combo Chart Example";

            // Save the workbook as a macro‑enabled XLSX file (.xlsm)
            workbook.Save("ComboChart.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
