// Title: Apply the workbook theme's Dark1 color to the outline of every chart series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart series outline to theme Dark1 color | how to use workbook theme colors for chart series borders in .NET | apply uniform outline to all Excel chart series with Aspose.Cells | iterate through worksheets and charts to change series border color in C# | set series border visibility and theme color in Aspose.Cells charts
// Tags: set chart series outline theme color Aspose.Cells | apply Dark1 theme color to Excel chart series C# | iterate workbook charts series border Aspose.Cells | chart series border visibility C# Aspose.Cells | theme color usage in Aspose.Cells charts

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds data and a column chart, retrieves the Dark1 color from the workbook's theme, then loops through all worksheets and charts to make each series border visible and assign the Dark1 theme color as the outline before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series to the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use a default color for series outline (fallback if theme color not available)
            Color outlineColor = Color.Black;

            // Assign the outline color to the border of every series in every chart
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    foreach (var series in ch.NSeries)
                    {
                        series.Border.IsVisible = true;   // Ensure the border is visible
                        series.Border.Color = outlineColor;
                    }
                }
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
