// Title: Generate an Excel workbook with a worksheet and a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to create a new workbook, add a worksheet named "DataSheet", fill cells A1:B5 with sample data, and insert a column chart that references the populated range. | Show how to configure the NSeries data range and category labels for a column chart in Aspose.Cells, and set a custom chart title. | Demonstrate saving the workbook to a specific file path after the chart is added, creating the output directory if it does not exist.
// Common Searches: Aspose.Cells C# create column chart from range B2:B5 with categories A2:A5 | how to add a chart to a worksheet using Aspose.Cells for .NET | saving an Aspose.Cells workbook to a custom folder in C# | set chart title and position for column chart in Aspose.Cells | programmatically generate Excel file with monthly sales chart using Aspose.Cells
// Tags: Aspose.Cells create workbook with column chart C# | Aspose.Cells set NSeries data range column chart | Aspose.Cells add worksheet and populate data for chart | Aspose.Cells save Excel file to custom path | Aspose.Cells configure chart title and position

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds a worksheet named DataSheet, populates cells A1:B5 with monthly sales data, inserts a column chart positioned between rows 7‑20 and columns A‑H, assigns the data and category ranges, sets a chart title, ensures the output directory exists, and saves the file as ColumnChartWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (added by default)
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DataSheet";

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");
                worksheet.Cells["A5"].PutValue("Apr");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["B4"].PutValue(180);
                worksheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet
                int chartUpperLeftRow = 7;
                int chartUpperLeftColumn = 0;
                int chartLowerRightRow = 20;
                int chartLowerRightColumn = 7;

                // Add returns the index of the newly created chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, chartUpperLeftRow, chartUpperLeftColumn, chartLowerRightRow, chartLowerRightColumn);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart series (values in column B)
                chart.NSeries.Add("B2:B5", false);

                // Set the category axis labels (categories in column A)
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: set chart title
                chart.Title.Text = "Monthly Sales";

                // Determine output path and ensure directory exists
                string outputFile = "ColumnChartWorkbook.xlsx";
                string outputPath = Path.GetFullPath(outputFile);
                string outputDir = Path.GetDirectoryName(outputPath);

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
