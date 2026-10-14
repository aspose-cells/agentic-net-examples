// Title: Create a multi‑series column chart in Aspose.Cells for .NET and freeze the data rows above the chart
// AI Prompts: Write a C# program using Aspose.Cells that builds a worksheet with category labels and three numeric series, adds a column chart referencing those series, sets the category axis to the first column, freezes the rows containing the source data, and saves the workbook as an .xlsx file. | Using Aspose.Cells for .NET, generate a column chart from ranges B2:B6, C2:C6, D2:D6, assign A2:A6 as category labels, apply FreezePanes at row 7 to keep the data visible while scrolling, and export the file.
// Common Searches: how to add multiple series to a column chart with Aspose.Cells in C# | Aspose.Cells freeze panes above a chart programmatically | set category axis range for a chart in Aspose.Cells .NET | create multi series column chart and keep source rows visible in Excel using Aspose.Cells | C# Aspose.Cells example for chart with three data series and frozen rows
// Tags: column chart multiple series Aspose.Cells | freeze panes above chart Aspose.Cells | category axis range definition Aspose.Cells | populate worksheet series data C# Aspose.Cells | export workbook with multi‑series chart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a worksheet named 'Data', writes category labels and three numeric series, builds a column chart that references those series, sets the X‑axis labels from the first column, freezes the rows containing the source data (rows 1‑6) to keep them visible while scrolling, and saves the file as MultiSeriesChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Write header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");

            // Sample data
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] series1 = { 10, 20, 30, 40, 50 };
            double[] series2 = { 15, 25, 35, 45, 55 };
            double[] series3 = { 12, 22, 32, 42, 52 };

            // Populate worksheet with data
            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Data starts at row 2 (1‑based)
                sheet.Cells[row, 0].PutValue(categories[i]); // Column A
                sheet.Cells[row, 1].PutValue(series1[i]);    // Column B
                sheet.Cells[row, 2].PutValue(series2[i]);    // Column C
                sheet.Cells[row, 3].PutValue(series3[i]);    // Column D
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Multi‑Series Column Chart";

            // Define series ranges
            chart.NSeries.Add("Data!$B$2:$B$6", true); // Series 1
            chart.NSeries.Add("Data!$C$2:$C$6", true); // Series 2
            chart.NSeries.Add("Data!$D$2:$D$6", true); // Series 3

            // Set category (X‑axis) labels
            chart.NSeries.CategoryData = "Data!$A$2:$A$6";

            // Freeze rows 1‑6 (rows above row 7)
            sheet.FreezePanes(7, 0, 0, 0);

            // Prepare output path
            string outputPath = "MultiSeriesChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            // Ensure the output directory exists
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
