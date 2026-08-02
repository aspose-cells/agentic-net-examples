using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Paths to the source workbooks
            string[] workbookFiles = new string[]
            {
                @"C:\Input\Book1.xlsx",
                @"C:\Input\Book2.xlsx",
                @"C:\Input\Book3.xlsx"
            };

            // Path to the chart template file (.crtx)
            string chartTemplatePath = @"C:\Templates\MyChartTemplate.crtx";

            // Verify that the chart template exists
            if (!File.Exists(chartTemplatePath))
            {
                Console.WriteLine($"Template file not found: {chartTemplatePath}");
                return;
            }

            // Load the template bytes once
            byte[] templateData = File.ReadAllBytes(chartTemplatePath);

            // Output folder for the modified workbooks
            string outputFolder = @"C:\Output\";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            foreach (string filePath in workbookFiles)
            {
                // Verify that the source workbook exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Workbook not found, skipping: {filePath}");
                    continue;
                }

                // Load the workbook from file
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (adjust as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Define the data range for the chart
                string dataRange = "A1:B5";

                // Add a chart using the preset template.
                // Parameters: template bytes, data range, isVertical, topRow, leftColumn, bottomRow, rightColumn
                int chartIndex = sheet.Charts.Add(
                    templateData,   // chart template bytes
                    dataRange,      // data range for the chart
                    true,           // plot series by column (vertical)
                    5,              // top row of the chart
                    0,              // left column of the chart
                    20,             // bottom row of the chart
                    8);             // right column of the chart

                // Optional further customization
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Chart with Preset Template";

                // Build the output file name
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string outputPath = Path.Combine(outputFolder, fileName + "_WithChart.xlsx");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed and saved: {outputPath}");
            }

            Console.WriteLine("All workbooks processed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}