using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplateDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the existing chart template file (.crtx)
                string templatePath = "ChartTemplate.crtx";

                // Load the template file into a byte array if it exists
                byte[] templateData = null;
                bool templateExists = File.Exists(templatePath);
                if (templateExists)
                {
                    templateData = File.ReadAllBytes(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file not found: {templatePath}. A default chart will be created.");
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                int chartIndex;

                if (templateExists && templateData != null)
                {
                    // Add a new chart using the preset template
                    chartIndex = sheet.Charts.Add(
                        templateData,          // template bytes
                        "A1:B4",               // data range
                        true,                  // series by column
                        5,                     // top row
                        0,                     // left column
                        20,                    // bottom row
                        8);                    // right column
                }
                else
                {
                    // Add a chart without a template (default style)
                    chartIndex = sheet.Charts.Add(
                        ChartType.Column,      // default chart type
                        5,                     // top row
                        0,                     // left column
                        20,                    // bottom row
                        8);                    // right column

                    // Set the data source manually
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";
                }

                // Optional further customization
                Chart addedChart = sheet.Charts[chartIndex];
                addedChart.Title.Text = "Styled Chart Using Template";

                // Save the workbook
                string outputPath = "ChartWithTemplate.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}