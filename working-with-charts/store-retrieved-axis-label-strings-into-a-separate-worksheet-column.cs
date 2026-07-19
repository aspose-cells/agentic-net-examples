// Title: Save Aspose.Cells Category Axis Labels to a Worksheet Column (C# .NET)
// Description: This example creates a workbook, adds sample data, inserts a column chart, runs Chart.Calculate to generate axis values, extracts the category‑axis texts with CategoryAxis.GetAxisTexts(), and writes each label into column D (with a header) before saving the file.
// Keywords: Aspose.Cells chart axis labels | CategoryAxis.GetAxisTexts C# | write axis texts to cells | export chart X‑axis values | Aspose.Cells .NET example
// Common Searches: retrieve category axis labels Aspose.Cells | store chart axis texts in Excel using C# | Aspose.Cells Chart.Calculate then GetAxisTexts | write chart X‑axis labels to worksheet column | Aspose.Cells export axis labels to cells
// Developer Intent: Extract the X‑axis (category) labels from an Aspose.Cells chart and place them into a worksheet column.
// Use Cases: Generate a separate table of axis labels for reporting dashboards. | Archive chart categories alongside raw data for audit trails. | Feed axis labels into downstream calculations or data‑validation rules.
// AI Prompts: Provide C# code that calculates an Aspose.Cells chart and saves its category axis labels in column D of the same sheet. | Show how to use Chart.Calculate and CategoryAxis.GetAxisTexts to export X‑axis labels to Excel cells with Aspose.Cells. | Explain the steps to retrieve chart axis texts in Aspose.Cells and write them to a worksheet range.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds sample data, inserts a column chart, runs Chart.Calculate to generate axis values, extracts the category‑axis texts with CategoryAxis.GetAxisTexts(), and writes each label into column D (with a header) before saving the file.
    public class StoreAxisLabelsInWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
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
                int startRow = 1;          // zero‑based index (row 2 in Excel)
                int targetColumn = 3;      // zero‑based index for column D

                for (int i = 0; i < axisLabels.Length; i++)
                {
                    worksheet.Cells[startRow + i, targetColumn].PutValue(axisLabels[i]);
                }

                // Add a header for the stored labels
                worksheet.Cells[0, targetColumn].PutValue("Axis Labels");

                // Save the workbook
                string outputPath = "AxisLabelsStored.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            StoreAxisLabelsInWorksheet.Run();
        }
    }
}
