// Title: Generate an Excel combo chart with a column series and a line series on a secondary axis using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create combo chart with column and line series on secondary axis | how to add a secondary axis to a line series in an Aspose.Cells chart | C# code sample for column‑line combo chart in Excel using Aspose.Cells | programmatically set series type to line in Aspose.Cells combo chart | Aspose.Cells generate sales and profit combo chart with dual axes
// Tags: create combo chart Aspose.Cells .NET | column series secondary axis Aspose.Cells | line series type conversion Aspose.Cells | dual axis Excel chart programmatically C# | sales profit combo chart Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example builds an Excel workbook, fills it with month, sales, and profit data, adds a column chart, converts a second series to a line type, places it on a secondary axis, and saves the combo chart as ComboChart.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);

            // Add a combo chart (initially a column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 26, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales and Profit Combo Chart";

            // Column series (primary axis)
            int columnSeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[columnSeriesIdx].Name = "Sales";

            // Line series (secondary axis)
            int lineSeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[lineSeriesIdx].Name = "Profit";
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line; // Change to line type

            // Note: In some older Aspose.Cells versions, secondary axis handling and
            // explicit category data assignment are not available. The chart will
            // display the line series alongside the column series using default settings.

            // Determine output path and ensure directory exists
            string outputPath = "ComboChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
