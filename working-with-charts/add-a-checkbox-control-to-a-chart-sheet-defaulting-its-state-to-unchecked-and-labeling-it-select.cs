// Title: Add an unchecked simulated checkbox labeled 'Select' to a chart sheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a rectangle shape on a chart sheet, sets its text to 'Select', and makes it appear as an unchecked checkbox with Aspose.Cells. | Show how to position a simulated checkbox shape beside a column chart in a workbook and save the result as an .xlsx file using Aspose.Cells for .NET.
// Common Searches: how to create a simulated checkbox on an Excel chart sheet with Aspose.Cells C# | Aspose.Cells add rectangle shape as checkbox on chart worksheet | C# Aspose.Cells default unchecked checkbox label Select | place UI shape next to chart in Aspose.Cells workbook
// Tags: add rectangle shape as checkbox Aspose.Cells | simulated form control on chart sheet C# | insert unchecked checkbox label Select Excel | Aspose.Cells chart worksheet shape positioning | save workbook with checkbox shape .xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a worksheet with a column chart, and simulates an unchecked checkbox by inserting a rectangle shape labeled 'Select' on the same sheet. The workbook is then saved as ChartWithCheckbox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add a new worksheet to host the chart and the simulated checkbox
                var worksheet = workbook.Worksheets.Add("SheetWithChart");

                // Add a simple column chart to the worksheet (positioned at rows 5‑15, columns A‑F)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                var chart = worksheet.Charts[chartIndex];
                chart.Title.Text = "Sample Chart";

                // Simulate a checkbox using a rectangle shape (Aspose.Cells older versions may lack form controls)
                // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
                var checkBoxShape = worksheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // use MsoDrawingType for compatibility
                    1,   // row index (0‑based)
                    1,   // column index (0‑based)
                    0,   // top offset in pixels
                    0,   // left offset in pixels
                    100, // width in pixels
                    20   // height in pixels
                );

                // Set the label of the simulated checkbox
                checkBoxShape.Text = "Select";

                // Determine output path and ensure directory exists
                string outputPath = "ChartWithCheckbox.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
