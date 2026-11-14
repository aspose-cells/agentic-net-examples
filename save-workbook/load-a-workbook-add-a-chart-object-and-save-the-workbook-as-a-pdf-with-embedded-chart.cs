// Title: Create a column chart in an Excel workbook and save it as a PDF with the chart embedded using Aspose.Cells for .NET
// Common Searches: c# aspose.cells add column chart to existing workbook and export to pdf | how to embed an Excel chart in a PDF when saving with Aspose.Cells | aspose.cells example for creating chart from cell range and converting to PDF | save workbook as pdf with chart hidden legend using aspose.cells c# | aspose.cells .net core generate column chart and output PDF file
// Tags: add column chart Aspose.Cells | chart data range definition Aspose.Cells | disable chart legend Aspose.Cells | convert workbook to PDF Aspose.Cells | chart rendering in PDF Aspose.Cells | c# Aspose.Cells chart generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel file, adds a column chart based on a specified range, sets a title, hides the legend, and then saves the workbook as a PDF where the chart is rendered inside the document.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet (rows 5‑20, columns 0‑10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sales Data";

            // Define the data range for the chart (A1:B5)
            int firstRow = 0;          // zero‑based index (A1)
            int firstColumn = 0;       // column A
            int totalRows = 5;         // rows 1‑5

            // Add a series using the values column (B)
            string valuesRange = sheet.Cells.CreateRange(firstRow, firstColumn + 1, totalRows, 1).RefersTo;
            chart.NSeries.Add(valuesRange, true);

            // Set the category (X) axis data using the categories column (A)
            string categoriesRange = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, 1).RefersTo;
            chart.NSeries.CategoryData = categoriesRange;

            // Optional: customize chart appearance
            chart.ShowLegend = false;

            // Save the workbook as PDF (chart will be embedded)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
