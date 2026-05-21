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
                // Paths of the workbooks to process
                string[] workbookFiles = new string[]
                {
                    "Workbook1.xlsx",
                    "Workbook2.xlsx",
                    "Workbook3.xlsx"
                };

                // Path to the predefined chart template (.crtx)
                string templatePath = "ChartTemplate.crtx";

                // Verify that the chart template exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the chart template into a byte array (required by Add method)
                byte[] templateData = File.ReadAllBytes(templatePath);

                // Iterate through each workbook
                foreach (string filePath in workbookFiles)
                {
                    // Verify that the workbook file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Workbook file not found: {filePath}");
                        continue; // Skip to next file
                    }

                    // Load the workbook from file
                    Workbook workbook = new Workbook(filePath);

                    // Access the first worksheet (adjust if needed)
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Define the data range for the chart (adjust to your data)
                    string dataRange = "A1:B5";

                    // Add a chart using the template data.
                    // Parameters: template bytes, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                    int chartIndex = worksheet.Charts.Add(
                        templateData,
                        dataRange,
                        true,   // Plot series by column
                        5,      // Top row of chart
                        0,      // Left column of chart
                        20,     // Bottom row of chart
                        8       // Right column of chart
                    );

                    // Optional: further customize the chart if needed
                    Chart chart = worksheet.Charts[chartIndex];
                    chart.Title.Text = $"Chart added to {Path.GetFileName(filePath)}";

                    // Save the modified workbook (overwrites the original file)
                    workbook.Save(filePath);
                }

                Console.WriteLine("All workbooks have been updated with the chart template.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}