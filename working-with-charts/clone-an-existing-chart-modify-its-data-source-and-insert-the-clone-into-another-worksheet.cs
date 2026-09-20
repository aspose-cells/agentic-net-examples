// Title: Clone an existing chart, change its data source, and insert the cloned chart into another worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that copies the first chart from Sheet1, adds a new chart of the same type to Sheet2 at row 5 column 1, and sets its series to use Sheet2!$C$2:$C$5 values and Sheet2!$B$2:$B$5 categories. | Generate a C# snippet that clones a chart object, clears its NSeries collection, assigns new category and value ranges from a different worksheet, and saves the workbook. | Create an Aspose.Cells example that loads a workbook, duplicates a chart onto another sheet, updates the data source to a new range, and specifies the chart position and size before saving.
// Common Searches: Aspose.Cells C# clone chart to another worksheet and set new data range | How to change the data source of a copied chart using Aspose.Cells for .NET | C# example for adding a chart with custom series after cloning with Aspose.Cells
// Tags: chart duplication Aspose.Cells C# | modify chart series range Aspose.Cells | add chart to another worksheet Aspose.Cells | set chart position and size Aspose.Cells | reset chart series Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// This C# program loads a workbook, clones the first chart from Sheet1, creates a new chart of the same type on Sheet2 at a specified location, clears its existing series, assigns new value and category ranges from Sheet2, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook that contains the original chart
            Workbook workbook = new Workbook(sourcePath);

            // Get the worksheet and the chart you want to clone
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // adjust name as needed
            if (sourceSheet == null || sourceSheet.Charts.Count == 0)
            {
                Console.WriteLine("Source worksheet or chart not found.");
                return;
            }

            Chart originalChart = sourceSheet.Charts[0]; // assumes the chart is the first one

            // Destination worksheet
            Worksheet targetSheet = workbook.Worksheets["Sheet2"]; // adjust name as needed
            if (targetSheet == null)
            {
                Console.WriteLine("Target worksheet not found.");
                return;
            }

            // Position and size for the cloned chart
            int upperLeftRow = 5;      // zero‑based row index
            int upperLeftColumn = 1;   // zero‑based column index
            int height = 400;          // height in pixels
            int width = 600;           // width in pixels

            // Add a new chart on the target sheet with the same type as the original
            int chartIndex = targetSheet.Charts.Add(originalChart.Type, upperLeftRow, upperLeftColumn, height, width);
            Chart clonedChart = targetSheet.Charts[chartIndex];

            // Modify the data source of the cloned chart
            // Example: use data from Sheet2 range B2:B5 for categories and C2:C5 for values
            clonedChart.NSeries.Clear(); // remove existing series
            clonedChart.NSeries.Add("Sheet2!$C$2:$C$5", true); // values series
            clonedChart.NSeries.CategoryData = "Sheet2!$B$2:$B$5"; // categories

            // Save the workbook with the new chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
