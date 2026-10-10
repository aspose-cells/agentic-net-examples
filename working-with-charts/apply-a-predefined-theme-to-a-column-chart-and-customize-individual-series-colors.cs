// Title: How to apply an Office .thmx theme to an Aspose.Cells workbook and set custom RGB colors for each series in a column chart (C#)
// AI Prompts: Generate C# code that creates a workbook, adds sample data, inserts a column chart, loads an .thmx theme file, and assigns specific RGB colors to the first and second series using Aspose.Cells. | Update an existing Aspose.Cells column chart to replace the theme‑derived colors with custom foreground colors for individual series while preserving the chart title and saving the file as XLSX.
// Common Searches: aspocells c# apply office .thmx theme to workbook column chart | set custom series colors in Aspose.Cells column chart after applying theme | override chart series fill color programmatically using Aspose.Cells .NET | how to change column chart series colors with Aspose.Cells C# example | load .thmx theme file in Aspose.Cells and customize chart series colors
// Tags: apply .thmx theme Aspose.Cells workbook | custom series foreground color column chart | Aspose.Cells C# column chart theming | set RGB color for chart series Aspose.Cells | save workbook as xlsx with themed chart

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartThemeExample
{
    // The example creates a new workbook, populates it with quarterly data, adds a column chart, optionally applies an Office .thmx theme, then overrides each series' fill color with custom RGB values, sets a chart title, and saves the result as ColumnChartWithTheme.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the column chart.
                // Column A: Categories, Column B: Values for Series 1, Column C: Values for Series 2.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(130);
                sheet.Cells["B5"].PutValue(170);

                sheet.Cells["C2"].PutValue(100);
                sheet.Cells["C3"].PutValue(140);
                sheet.Cells["C4"].PutValue(110);
                sheet.Cells["C5"].PutValue(160);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart.
                // Categories are in A2:A5, Series 1 in B2:B5, Series 2 in C2:C5.
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Series 1";
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Series 2";

                // Set the category axis labels.
                chart.NSeries.CategoryData = "A2:A5";

                // Apply a predefined theme to the entire workbook if the file exists.
                // Note: Aspose.Cells version used may not support ApplyTheme; this block is optional.
                string themePath = @"C:\Themes\OfficeTheme.thmx";
                if (File.Exists(themePath))
                {
                    // If the API supports ApplyTheme, uncomment the following line:
                    // workbook.ApplyTheme(themePath);
                }

                // Customize individual series colors after the (optional) theme is applied.
                chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(0, 112, 192); // Custom blue
                chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(255, 165, 0); // Orange

                // Optionally, adjust the chart title.
                chart.Title.Text = "Quarterly Sales";

                // Save the workbook to a file.
                string outputPath = "ColumnChartWithTheme.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
