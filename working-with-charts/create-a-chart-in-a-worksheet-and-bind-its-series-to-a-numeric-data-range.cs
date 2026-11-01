// Title: Create a column chart in Aspose.Cells (C#) and bind its series to cells B2:B6
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart to the first worksheet, sets the chart title to "Monthly Sales", binds the series to the range B2:B6, assigns the series name "Sales", and saves the workbook. | Show how to programmatically bind a chart series to a numeric cell range in Aspose.Cells for .NET and customize the series label. | Generate an example that positions a column chart from row 7, column 0 to row 25, column 7 using Aspose.Cells C#.
// Common Searches: aspnet add column chart using Aspose.Cells and link series data to cells B2 through B6 | c# Aspose.Cells example for connecting chart series with numeric cells | programmatically set chart title and series label in Aspose.Cells .NET | position a column chart from row 7 column 0 to row 25 column 7 with Aspose.Cells | save Excel workbook after creating chart using Aspose.Cells C#
// Tags: add column chart Aspose.Cells C# | bind chart series to cell range Aspose.Cells | set chart title Aspose.Cells | name chart series Aspose.Cells | save workbook after chart creation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A1:B6 with category labels and numeric values, adds a column chart positioned from row 7, column 0 to row 25, column 7, sets the chart title to "Monthly Sales", binds the series to the range B2:B6, names the series "Sales", and saves the file as ChartExample.xlsx.
class ChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["A6"].PutValue("May");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(25);
            sheet.Cells["B6"].PutValue(15);

            // Add a column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title.
            chart.Title.Text = "Monthly Sales";

            // Bind series data.
            chart.NSeries.Add("B2:B6", true);
            // Category data can be inferred; explicit setting removed for compatibility.
            chart.NSeries[0].Name = "Sales";

            // Save the workbook.
            string outputPath = "ChartExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
