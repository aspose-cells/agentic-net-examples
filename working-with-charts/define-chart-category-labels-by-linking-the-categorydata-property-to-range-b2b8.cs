// Title: How to assign month names as chart category labels from cells B2:B8 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that creates a new workbook, fills column B with month names and column C with sales figures, adds a column chart, and sets the chart's CategoryData to the range B2:B8 using Aspose.Cells. | Write code that binds the CategoryData property of an Aspose.Cells chart to a worksheet range, links series values to C2:C8, sets a chart title, and saves the workbook as an .xlsx file. | Produce a snippet that populates columns B and C, creates a column chart, assigns "Sales by Month" as the title, and links the chart's category labels to the month column (B2:B8).
// Common Searches: Aspose.Cells C# set chart category labels from a worksheet range | How to bind CategoryData property to cells B2:B8 in an Aspose.Cells chart | Create column chart with month names as categories using Aspose.Cells for .NET | C# example linking chart categories to column B in an Aspose.Cells workbook | Aspose.Cells chart category data range B2:B8 tutorial
// Tags: Aspose.Cells chart category data binding | C# set chart CategoryData range | Aspose.Cells column chart from worksheet data | save workbook with chart Aspose.Cells | populate worksheet cells for chart labels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes month names to column B and sales values to column C, adds a column chart, links the series to C2:C8, sets the chart's CategoryData to B2:B8, assigns the title "Sales by Month", and saves the file as ChartWithCategoryLabels.xlsx.
class ChartCategoryLabelsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the chart
            double[] seriesValues = { 10, 20, 30, 40, 50, 60, 70 };
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul" };

            for (int i = 0; i < seriesValues.Length; i++)
            {
                sheet.Cells[i + 1, 1].PutValue(categories[i]);   // Column B (index 1)
                sheet.Cells[i + 1, 2].PutValue(seriesValues[i]); // Column C (index 2)
            }

            // Add a column chart (rows 5‑15, columns 0‑5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Link series values and category labels
            chart.NSeries.Add("C2:C8", true);
            chart.NSeries.CategoryData = "B2:B8";

            // Set chart title
            chart.Title.Text = "Sales by Month";

            // Ensure the output directory exists
            string outputPath = "ChartWithCategoryLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
