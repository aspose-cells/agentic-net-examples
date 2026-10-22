// Title: How to add a column chart to the first worksheet of an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing Excel workbook (or creates a new one), fills cells A1:B4 with sample data, inserts a column chart on the first sheet, sets the chart title, defines the series and category ranges, and saves the file. | Demonstrate how to place a column chart at a specific cell rectangle (e.g., rows 5‑20, columns A‑F) and customize its data source using the Aspose.Cells Chart API in C#. | Provide an example that calculates the last used row dynamically and creates a column chart that automatically references the entire data range in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add column chart to first worksheet example | set series and category data for column chart using Aspose.Cells .NET | position column chart at rows 5 to 20 columns A to F with Aspose.Cells | load or create workbook then insert chart Aspose.Cells C#
// Tags: Aspose.Cells add column chart C# | column chart positioning Aspose.Cells | set chart series Aspose.Cells API | save workbook with chart Aspose.Cells .NET | populate sample data Excel Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing workbook or creates a new one, writes sample data to cells A1:B4, adds a column chart to the first worksheet positioned from row 5 to row 20 and column A to column F, sets the chart title, assigns series (B2:B4) and category (A2:A4) ranges, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sample Column Chart";

            // Define the series data and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
