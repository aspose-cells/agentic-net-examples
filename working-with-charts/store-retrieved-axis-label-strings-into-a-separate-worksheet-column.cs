// Title: Aspose.Cells for .NET – Store Chart Category Axis Labels in a Worksheet Column
// Description: Creates a workbook, adds sample data, inserts a column chart, calculates it to generate axis texts, extracts the category axis labels with GetAxisTexts(), and writes each label into column D (with a header) before saving the file.
// Keywords: Aspose.Cells C# chart axis labels | GetAxisTexts example | store category axis texts in Excel cells | write chart labels to worksheet column | Aspose.Cells CategoryAxis extraction | export chart axis values .NET | column chart label retrieval
// Common Searches: Aspose.Cells retrieve category axis labels | C# write chart axis texts to Excel column | GetAxisTexts Aspose.Cells example | store chart X‑axis values in worksheet | Aspose.Cells chart label extraction tutorial
// Developer Intent: Extract the category (X‑axis) labels from an Aspose.Cells chart and write them into a specified worksheet column.
// Use Cases: Generate a reference table of chart categories for downstream formulas or reporting. | Synchronize chart labels with raw data when building dynamic dashboards. | Export axis labels for integration with external analytics tools or APIs.
// AI Prompts: Provide C# code that extracts both category and value axis texts from an Aspose.Cells chart and saves them to separate worksheet columns. | Show how to handle multi‑series charts and store each series' axis labels in distinct columns using Aspose.Cells. | Suggest robust error‑handling patterns for Chart.Calculate and GetAxisTexts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, calculates it to generate axis texts, extracts the category axis labels with GetAxisTexts(), and writes each label into column D (with a header) before saving the file.
    public class StoreAxisLabelsInWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart (categories in column A, values in column B)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and the category (X) axis
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Calculate the chart so that axis labels are generated
                chart.Calculate();

                // Retrieve the category axis labels
                string[] axisLabels = chart.CategoryAxis.GetAxisTexts();

                // Store the retrieved labels into column D, starting from row 2
                int startRow = 1; // zero‑based index (row 2 in Excel)
                int targetColumn = 3; // zero‑based index for column D

                for (int i = 0; i < axisLabels.Length; i++)
                {
                    worksheet.Cells[startRow + i, targetColumn].PutValue(axisLabels[i]);
                }

                // Add a header for the stored labels
                worksheet.Cells[0, targetColumn].PutValue("Axis Labels");

                // Save the workbook
                string outputPath = "AxisLabelsStored.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StoreAxisLabelsInWorksheet.Run();
        }
    }
}
