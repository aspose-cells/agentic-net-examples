// Title: Format Z‑axis (Series axis) labels as percentages in a 3‑D column chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set Z axis tick label format to percent in 3D column chart | how to display series axis values as percentages using Aspose.Cells .NET | C# Aspose.Cells change depth axis number format to 0% for 3D charts | apply percentage number format to chart Z axis in Aspose.Cells workbook | format 3D column chart series axis labels as percent with Aspose.Cells
// Tags: set series axis number format Aspose.Cells | 3D column chart Z axis percentage formatting | Aspose.Cells C# chart tick label number format | apply percentage format to chart axis labels | Aspose.Cells series axis formatting example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a 3‑D column chart with two data series, accesses the chart's SeriesAxis (Z‑axis), and sets its TickLabels.NumberFormat to "0%" so the axis labels appear as percentages. The workbook is then saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a 3‑D chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(0.10);
            sheet.Cells["B3"].PutValue(0.25);
            sheet.Cells["B4"].PutValue(0.40);
            sheet.Cells["B5"].PutValue(0.15);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(0.20);
            sheet.Cells["C3"].PutValue(0.30);
            sheet.Cells["C4"].PutValue(0.35);
            sheet.Cells["C5"].PutValue(0.25);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Series1";
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Series2";

            // Set category (X) data
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[1].DataLabels.ShowValue = true;

            // ------------------------------------------------------------
            // Set Z‑axis (Series axis) label number format to percentage
            // ------------------------------------------------------------
            Axis zAxis = chart.SeriesAxis; // Depth axis for 3‑D charts
            if (zAxis != null)
            {
                // Apply percentage format to the tick labels
                zAxis.TickLabels.NumberFormat = "0%";
                // No need to set IsNumberFormatLinked (property does not exist)
            }
            else
            {
                Console.WriteLine("Series (Z) axis not found on the chart.");
            }

            // Save the workbook
            string outputPath = "ZAxisPercentageChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
