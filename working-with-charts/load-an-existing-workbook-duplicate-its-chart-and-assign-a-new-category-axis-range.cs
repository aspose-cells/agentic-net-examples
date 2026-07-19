// Title: Copy a chart to a new workbook and set a new category axis range with Aspose.Cells for .NET
// Description: Loads a source workbook, copies its first worksheet (including charts) into a new workbook, accesses the duplicated chart, updates the CategoryData of its NSeries to a new range, and saves the result. Includes file‑existence checks and error handling.
// Keywords: Aspose.Cells chart copy | C# duplicate chart workbook | set CategoryData Aspose.Cells | copy worksheet with charts .NET | modify chart data source programmatically
// Common Searches: Aspose.Cells copy chart to another workbook | change category axis range after chart copy | C# duplicate worksheet chart Aspose.Cells | update NSeries.CategoryData in copied chart | how to clone a chart with Aspose.Cells
// Developer Intent: Duplicate a chart from an existing workbook into a new workbook and assign a different category axis range.
// Use Cases: Reuse a chart layout across multiple reports while pointing each to its own data block. | Generate a series of workbooks where each contains a copied chart that reflects a distinct category range. | Automate the refresh of chart categories after copying a template worksheet.
// AI Prompts: Generate C# code using Aspose.Cells that copies the first worksheet with charts to a new workbook and changes the first chart's CategoryData to "B2:B7". | Explain best practices for safely cloning a chart and updating its data source, including handling worksheets that contain no charts. | Provide a step‑by‑step tutorial for copying a chart, setting a new category axis range, and saving the workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a source workbook, copies its first worksheet (including charts) into a new workbook, accesses the duplicated chart, updates the CategoryData of its NSeries to a new range, and saves the result. Includes file‑existence checks and error handling.
class DuplicateChartExample
{
    static void Main()
    {
        const string sourcePath = "SourceWorkbook.xlsx";
        const string destinationPath = "DestinationWorkbook.xlsx";

        try
        {
            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook that contains the original chart
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) workbook that will hold the duplicated chart
            Workbook destinationWorkbook = new Workbook();

            // Copy the first worksheet (which includes the chart) from the source workbook
            // Use the overload that accepts the worksheet name
            int copiedSheetIndex = destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0].Name);
            Worksheet copiedWorksheet = destinationWorkbook.Worksheets[copiedSheetIndex];

            // Ensure the copied worksheet contains at least one chart
            if (copiedWorksheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts were found in the copied worksheet.");

            // Access the duplicated chart in the copied worksheet
            Chart duplicatedChart = copiedWorksheet.Charts[0];

            // Assign a new category axis range to the duplicated chart
            // Adjust the range as needed for your data layout
            duplicatedChart.NSeries.CategoryData = "A10:A15";

            // Save the destination workbook with the duplicated chart
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to '{destinationPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
