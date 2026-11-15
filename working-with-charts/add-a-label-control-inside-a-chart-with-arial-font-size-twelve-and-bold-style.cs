// Title: Add a bold Arial label inside a chart title using Aspose.Cells for .NET (C#)
// Common Searches: how to set chart title font to Arial bold size 12 in Aspose.Cells C# | Aspose.Cells add custom label inside an Excel chart programmatically | C# example for formatting chart title with Arial bold using Aspose.Cells | change Excel chart title style to Arial 12 bold with Aspose.Cells .NET | Aspose.Cells chart title styling code snippet for .xlsx files
// Tags: Aspose.Cells chart title font styling | C# set chart title Arial bold with Aspose.Cells | Excel chart label formatting using Aspose.Cells | Aspose.Cells add label control to chart | Aspose.Cells .NET chart title customization | Aspose.Cells chart title size twelve

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Creates a workbook, adds a column chart with sample data, sets the chart title text to "Sales Overview" and applies Arial font, size 12, bold style, then saves the file as ChartWithLabel.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart (positioned from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Define series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure chart title
            chart.Title.Text = "Sales Overview";
            chart.Title.Font.Name = "Arial";
            chart.Title.Font.Size = 12;
            chart.Title.Font.IsBold = true;
            // Overlay property is not available in some versions; omitted for compatibility

            // Determine output path and ensure directory exists
            string outputPath = "ChartWithLabel.xlsx";
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
