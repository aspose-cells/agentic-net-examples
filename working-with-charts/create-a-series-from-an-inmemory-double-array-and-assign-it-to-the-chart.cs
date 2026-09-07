// Title: Create a column chart from an in‑memory double array and string labels with Aspose.Cells for .NET
// AI Prompts: Write C# code that imports a double[] into a worksheet, adds a column chart, and sets the series Y‑values to that range using Aspose.Cells. | Show how to import a string[] as category labels and assign it to the chart’s X‑axis in Aspose.Cells. | Modify the example to generate a line chart instead of a column chart while still using in‑memory arrays for data and labels.
// Common Searches: Aspose.Cells C# create column chart from double array without using a worksheet file | how to bind string array as X axis categories in Aspose.Cells chart | importing in‑memory numeric data into Excel and charting with Aspose.Cells .NET | save Excel workbook with chart built from arrays using Aspose.Cells | C# Aspose.Cells chart series from cell range generated from array
// Tags: import double array to worksheet Aspose.Cells | add column chart series from cell range C# | bind category axis to string array Aspose.Cells | save workbook with chart as xlsx Aspose.Cells | chart from in‑memory data C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsInMemoryArraySeries
{
    // The program creates a new workbook, imports a double[] into column A and a string[] into column B, adds a column chart whose Y‑values come from the double array and X‑axis categories from the string array, then saves the file as InMemoryArraySeries.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // In‑memory double array that will become the Y‑values of the chart series
                double[] yValues = new double[] { 10.5, 20.75, 30.0, 40.25 };

                // Import the double array into column A (0‑based index) vertically starting at cell A1
                sheet.Cells.ImportArray(yValues, 0, 0, true);

                // Optional: create category (X) labels in column B for better readability
                string[] xLabels = new string[] { "Q1", "Q2", "Q3", "Q4" };
                sheet.Cells.ImportArray(xLabels, 0, 1, true);

                // Add a column chart to the worksheet (positioned from row 5, column 0 to row 15, column 5)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Add a series whose Y‑values are taken from the imported double array (A1:A4)
                chart.NSeries.Add("=Sheet1!$A$1:$A$4", true);

                // Assign the X‑axis (category) data to the imported string labels (B1:B4)
                chart.NSeries.CategoryData = "=Sheet1!$B$1:$B$4";

                // Define output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "InMemoryArraySeries.xlsx");

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
