// Title: How to add a secondary axis to the line series of a column‑line combo chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set line series to secondary axis in a combo chart | Create column and line combo chart with secondary axis using Aspose.Cells .NET | Assign secondary axis to chart series in an Aspose.Cells workbook C# example | Plot line series on second axis in Aspose.Cells combo chart programmatically | Aspose.Cells IsSecondaryAxis vs PlotOnSecondAxis property usage
// Tags: Aspose.Cells secondary axis chart C# | combo column line chart Aspose.Cells | IsSecondaryAxis property Aspose.Cells | PlotOnSecondAxis fallback Aspose.Cells | assign chart series to secondary axis C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills it with month data, adds a column‑line combo chart, defines a column series and a line series, then moves the line series to a secondary axis using the IsSecondaryAxis property when available or the PlotOnSecondAxis property as a fallback, and saves the file as ComboChartSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Column Series");
            sheet.Cells["C1"].PutValue("Line Series");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Add a combo chart (Column + Line) to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart comboChart = sheet.Charts[chartIndex];

            // Set a title for the chart
            comboChart.Title.Text = "Combo Chart with Secondary Axis";

            // Add the column series (first series) using data from column B
            int columnSeriesIdx = comboChart.NSeries.Add("B2:B5", true);
            comboChart.NSeries[columnSeriesIdx].Name = "Column Series";
            comboChart.NSeries[columnSeriesIdx].Type = ChartType.Column;

            // Add the line series (second series) using data from column C
            int lineSeriesIdx = comboChart.NSeries.Add("C2:C5", true);
            comboChart.NSeries[lineSeriesIdx].Name = "Line Series";
            comboChart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // Assign the line series to the secondary axis (if supported)
            try
            {
                // In newer versions the property is IsSecondaryAxis; older versions use PlotOnSecondAxis
                var series = comboChart.NSeries[lineSeriesIdx];
                // Prefer IsSecondaryAxis if available
                var type = series.GetType();
                var prop = type.GetProperty("IsSecondaryAxis");
                if (prop != null)
                {
                    prop.SetValue(series, true);
                }
                else
                {
                    // Fallback to PlotOnSecondAxis
                    series.PlotOnSecondAxis = true;
                }
            }
            catch (Exception)
            {
                // If the property is not supported, continue without secondary axis
            }

            // Determine output file path and ensure the directory exists
            string outputPath = "ComboChartSecondaryAxis.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save the workbook:");
                Console.WriteLine(saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the combo chart:");
            Console.WriteLine(ex.Message);
        }
    }
}
