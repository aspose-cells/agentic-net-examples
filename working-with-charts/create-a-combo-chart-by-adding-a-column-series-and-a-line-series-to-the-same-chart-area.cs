// Title: How to create a combo chart with column and line series in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a column series and a line series to the same chart area in a worksheet. | Demonstrate setting the first series type to Column and the second series type to Line in an Aspose.Cells combo chart. | Show how to save the workbook containing the combo chart to a specific file path and create the output folder if it does not exist.
// Common Searches: Aspose.Cells C# create combo chart with column and line series | add mixed column and line series to one chart using Aspose.Cells .NET | programmatically generate combo chart in Excel with Aspose.Cells example | set individual series chart type in Aspose.Cells chart C# | save workbook with combo chart Aspose.Cells .NET
// Tags: Aspose.Cells create combo chart | combo chart column series Aspose.Cells | combo chart line series Aspose.Cells | set series type Aspose.Cells chart | save workbook Aspose.Cells Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a new workbook, populates sample data, adds a combo chart with a column series and a line series, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
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

            // Populate sample data.
            sheet.Cells["A1"].PutValue("Category");
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

            // Add a combo chart.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Combo Chart Example";

            // First series (Column).
            int seriesIndex1 = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[seriesIndex1].Type = ChartType.Column;
            chart.NSeries[seriesIndex1].Name = "Column Series";

            // Second series (Line).
            int seriesIndex2 = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[seriesIndex2].Type = ChartType.Line;
            chart.NSeries[seriesIndex2].Name = "Line Series";

            // Note: CategoryData property is not required for basic charts; categories will be taken from the first column.

            // Determine output file path.
            string outputPath = "ComboChart.xlsx";

            // Ensure the directory exists if a directory part is present.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
