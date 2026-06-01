using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendTagDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Store a custom tag on the chart using the Title (no direct Tag property)
                chart.Title.Text = "MyCustomLegendTag";

                // Save the workbook with the chart
                string filePath = "ChartWithLegendTag.xlsx";
                workbook.Save(filePath);

                // ----- Later retrieval of the custom tag -----
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Chart loadedChart = loadedSheet.Charts[0];

                // Retrieve the tag from the chart (stored in Title.Text)
                string retrievedTag = loadedChart.Title.Text;

                // Output the retrieved tag to verify
                Console.WriteLine("Retrieved Legend Tag: " + retrievedTag);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}