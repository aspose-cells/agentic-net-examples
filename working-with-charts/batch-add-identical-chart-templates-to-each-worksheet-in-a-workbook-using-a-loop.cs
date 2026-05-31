using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartBatch
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputWorkbookPath = "InputWorkbook.xlsx";
                const string chartTemplatePath = "ChartTemplate.crtx";
                const string outputWorkbookPath = "OutputWorkbook.xlsx";

                // Verify the input workbook exists
                if (!File.Exists(inputWorkbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputWorkbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Verify the chart template exists
                if (!File.Exists(chartTemplatePath))
                {
                    Console.WriteLine($"Chart template not found: {chartTemplatePath}");
                    return;
                }

                // Load the chart template into a byte array
                byte[] chartTemplateData = File.ReadAllBytes(chartTemplatePath);

                // Define the data range for the charts
                string dataRange = "A1:B4";

                // Add a chart to each worksheet using the template
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int chartIndex = sheet.Charts.Add(
                        chartTemplateData, // template data
                        dataRange,         // data range
                        true,              // plot series by column (vertical)
                        0,                 // top row
                        0,                 // left column
                        15,                // bottom row
                        8);                // right column

                    Chart chart = sheet.Charts[chartIndex];
                    chart.Title.Text = $"Chart on {sheet.Name}";
                }

                // Save the modified workbook
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved to {outputWorkbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}