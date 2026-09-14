// Title: Populate worksheet cells with source data and generate a column chart in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that fills cells A1:C4 with headers and numeric values, adds a column chart referencing that range, sets the chart title, and saves the workbook as an .xlsx file using Aspose.Cells. | Show how to bind a column chart to a data range by column, including headers, with Aspose.Cells Chart.SetChartDataRange in C#. | Demonstrate positioning a column chart on rows 5‑15 and columns 0‑5 after loading source data into a worksheet using Aspose.Cells. | Provide a snippet that outputs the generated Excel file to the current working directory and logs the saved path.
// Common Searches: Aspose.Cells C# fill worksheet cells and create column chart with data range by column | How to set chart data source including headers in Aspose.Cells .NET | Example of adding a column chart to an Excel sheet using Aspose.Cells in C# | Save generated Excel workbook with chart to current directory using Aspose.Cells
// Tags: fill worksheet with chart source data Aspose.Cells C# | bind column chart to data range by column Aspose.Cells | position column chart on worksheet Aspose.Cells | set chart title Aspose.Cells | save workbook as .xlsx file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataSource
{
    // The example creates a new workbook, writes headers and sample numeric data to cells A1:C4, adds a column chart positioned on the sheet, binds the chart to the data range by column (including headers), sets a custom chart title, and saves the file as ChartWithDataSource.xlsx in the current directory.
    public class PopulateDataForChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header cells
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["C1"].PutValue("Series2");

                // Populate sample data rows
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["C2"].PutValue(20);

                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["C3"].PutValue(40);

                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(50);
                worksheet.Cells["C4"].PutValue(60);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart (including headers)
                // The 'true' flag indicates that series are plotted by column
                chart.SetChartDataRange("A1:C4", true);

                // Optionally, set a chart title
                chart.Title.Text = "Sample Column Chart";

                // Determine output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithDataSource.xlsx");

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PopulateDataForChart.Run();
        }
    }
}
