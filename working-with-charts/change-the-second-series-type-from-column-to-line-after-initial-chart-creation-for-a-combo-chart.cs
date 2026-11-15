// Title: Change the second data series from column to line in an Aspose.Cells combo chart using C#
// Common Searches: C# Aspose.Cells how to convert a combo chart series from column to line | Aspose.Cells change second series chart type after creating combo chart | modify series type in Excel combo chart with Aspose.Cells .NET | set line series for second data series in Aspose.Cells combo chart example | Aspose.Cells programmatically switch column series to line in a combo chart
// Tags: Aspose.Cells combo chart series type change | C# Aspose.Cells column to line series conversion | modify Excel chart series Aspose.Cells .NET | dynamic chart series type Aspose.Cells | Excel combo chart customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds month labels and two data series, inserts a combo chart with both series initially as columns, then changes the second series to a line type, sets titles, and saves the file as an XLSX workbook.
class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data.
            // Category labels.
            cells["A2"].PutValue("Jan");
            cells["A3"].PutValue("Feb");
            cells["A4"].PutValue("Mar");
            cells["A5"].PutValue("Apr");

            // First series (Column).
            cells["B2"].PutValue(120);
            cells["B3"].PutValue(150);
            cells["B4"].PutValue(170);
            cells["B5"].PutValue(130);

            // Second series (will be changed to Line).
            cells["C2"].PutValue(80);
            cells["C3"].PutValue(90);
            cells["C4"].PutValue(100);
            cells["C5"].PutValue(95);

            // Add a Combo chart (initially both series as Column).
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (Column).
            chart.NSeries.Add("B2:B5", true);
            // Add the second series (Column initially).
            chart.NSeries.Add("C2:C5", true);

            // Change the second series type from Column to Line.
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: set chart title and axis titles.
            chart.Title.Text = "Sales Comparison";
            chart.CategoryAxis.Title.Text = "Month";
            chart.ValueAxis.Title.Text = "Amount";

            // Define output file path.
            string outputPath = "ComboChart_Output.xlsx";

            // Ensure the directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the combo chart:");
            Console.WriteLine(ex.Message);
        }
    }
}
