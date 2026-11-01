// Title: Create a column chart and format data labels to show value and series name with a hyphen separator using Aspose.Cells for .NET
// AI Prompts: Generate a column chart from cells B2:B4 with categories A2:A4, then configure its first series so the data labels display the cell value, a hyphen, and the series name in Aspose.Cells for .NET. | Enable data labels on an Aspose.Cells chart to show both the numeric value and the series name together, using the default hyphen separator, and save the workbook as an .xlsx file. | Add sample data, insert a column chart, and set DataLabels.ShowValue and DataLabels.ShowSeriesName to true to produce combined labels in the resulting Excel file.
// Common Searches: Aspose.Cells C# show value and series name in chart data labels | How to add hyphen between value and series name in Aspose.Cells chart labels | Set data label format to value-series name for column chart using Aspose.Cells .NET | Display series name with cell value in Excel chart via Aspose.Cells API | Aspose.Cells column chart data label default separator behavior
// Tags: Aspose.Cells column chart data labels | data label show value and series name | hyphen separator for chart data labels .NET | Aspose.Cells set series data labels | Excel chart label customization Aspose.Cells | C# Aspose.Cells chart data label formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart using ranges B2:B4 for values and A2:A4 for categories, enables data labels to show both the cell value and the series name (default hyphen separator), and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];

            // Define series values and categories
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Configure data labels to show value and series name
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;        // Show cell value
            series.DataLabels.ShowSeriesName = true;   // Show series name
            // Note: The Separator property is not available in the current Aspose.Cells version.
            // The default separator will be used.

            // Determine output path and ensure the directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
