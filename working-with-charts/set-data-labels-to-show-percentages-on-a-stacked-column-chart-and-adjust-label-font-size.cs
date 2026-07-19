// Title: Aspose.Cells for .NET – Add Percentage Labels & Adjust Font Size on a Stacked Column Chart (C#)
// Description: This C# example creates a workbook, fills it with sample data, inserts a stacked column chart, enables data labels that display each segment's percentage, and customizes the label font to 12 pt before saving the file as StackedColumnChart_PercentageLabels.xlsx.
// Keywords: Aspose.Cells C# stacked column chart | percentage data labels Aspose.Cells | chart label font size .NET | Excel chart customization Aspose | show percentages in column chart C# | Aspose.Cells chart series formatting | US developers Aspose.Cells examples | global Excel automation Aspose
// Common Searches: how to show percentages on a stacked column chart using Aspose.Cells | change data label font size in Aspose.Cells C# chart | Aspose.Cells add data labels to chart series | C# code for percentage labels in Excel stacked column chart | Aspose.Cells chart formatting tutorial
// Developer Intent: Add a stacked column chart, turn on percentage data labels, and set a custom font size for those labels using Aspose.Cells in C#.
// Use Cases: Quarterly sales dashboards that highlight each product's share as a percentage. | Financial reports where stacked columns need readable percentage markers. | Automated Excel workbook generation for business intelligence with styled chart labels.
// AI Prompts: Generate C# code with Aspose.Cells that creates a stacked column chart, shows percentage labels, and sets the label font to 12 pt. | Explain how to enable and style data labels (percentage, font size, color) for chart series in Aspose.Cells for .NET. | Provide step‑by‑step instructions to modify an existing Aspose.Cells chart to display segment percentages and customize label appearance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, fills it with sample data, inserts a stacked column chart, enables data labels that display each segment's percentage, and customizes the label font to 12 pt before saving the file as StackedColumnChart_PercentageLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a stacked column chart to the worksheet
            // Use ColumnStacked as the correct enum value for stacked column charts
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series data (two series) and category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Product A";
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Product B";
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels, show percentages, and set font size for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowPercentage = true;   // Display percentage values
                series.DataLabels.Font.Size = 12;          // Adjust label font size (e.g., 12 points)
            }

            // Determine output file path and ensure the directory exists
            string outputFile = "StackedColumnChart_PercentageLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
