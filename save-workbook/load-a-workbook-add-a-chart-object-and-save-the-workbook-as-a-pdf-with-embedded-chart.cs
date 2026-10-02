// Title: Create a column chart in an existing Excel workbook and export it as a PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file, inserts a column chart using A1:A5 for categories and B1:B5 for values, and saves the workbook as a PDF with the chart embedded via Aspose.Cells. | Adjust the chart bounds so the upper‑left corner starts at row 10, column 2 and the lower‑right corner ends at row 25, column 8, then export the workbook to PDF using Aspose.Cells. | Replace the column chart with a line chart, set a custom title "Sales Trend", enable the legend, and generate a PDF output with Aspose.Cells.
// Common Searches: how to add a column chart to an existing Excel file and convert it to PDF using Aspose.Cells in C# | Aspose.Cells C# export workbook to PDF with embedded chart from specific cell range | C# sample for inserting a chart into a loaded workbook and saving as PDF with Aspose.Cells | change chart dimensions in Aspose.Cells before saving workbook as PDF
// Tags: add column chart to worksheet Aspose.Cells | export workbook to PDF with chart Aspose.Cells | define chart placement cells Aspose.Cells | load existing workbook and create chart Aspose.Cells | embed chart in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // Loads 'input.xlsx', adds a column chart covering rows 5‑20 and columns A‑K using data from A1:A5 and B1:B5, sets a title and legend, then saves the workbook as 'output.pdf' where the chart is embedded.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define chart placement cells
            int upperLeftRow = 5;
            int upperLeftColumn = 0;
            int lowerRightRow = 20;
            int lowerRightColumn = 10;

            // Add a column chart to the worksheet (correct overload order)
            int chartIndex = sheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sample Chart";

            // Add a data series (categories from A1:A5, values from B1:B5)
            chart.NSeries.Add("A1:A5", true);          // categories
            chart.NSeries[0].Values = "B1:B5";         // values

            // Optional appearance settings
            chart.ShowLegend = true;

            // Save the workbook as PDF; the chart will be embedded in the PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
