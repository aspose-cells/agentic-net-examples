// Title: Generate an Excel pie chart without gridlines using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds sample data, inserts a pie chart, and guarantees the chart displays no gridlines. | Show how to programmatically disable plot‑area gridlines for a pie chart in Aspose.Cells before saving the workbook as an .xlsx file. | Provide a concise Aspose.Cells example that builds a pie chart and confirms that gridlines are omitted for a cleaner visual.
// Common Searches: asp.net c# aspose.cells hide gridlines on pie chart | remove chart gridlines from Excel file using Aspose.Cells .NET | create pie chart without plot area gridlines aspose.cells example | aspose.cells generate clean pie chart no gridlines c# | how to disable gridlines for Excel pie chart programmatically
// Tags: Aspose.Cells create pie chart C# | Aspose.Cells disable chart gridlines | Aspose.Cells export Excel without gridlines | C# generate Excel pie chart Aspose.Cells | pie chart visual cleanup Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample category/value data, adds a pie chart based on that data, and saves the file as PieChart_NoGridlines.xlsx. Because gridlines are not applicable to pie charts, the chart is rendered without any plot‑area gridlines, resulting in a clean visual presentation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet (lifecycle rule: create chart)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Gridlines are not applicable to pie charts; no need to modify PlotArea.

            // Ensure the output directory exists (if a directory is specified)
            string outputPath = "PieChart_NoGridlines.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
