// Title: How to verify an Excel data range for empty cells before generating a column chart with Aspose.Cells in C#
// AI Prompts: Iterate over a specified worksheet range, detect any cells whose Value is null, and abort chart creation when a blank cell is found using Aspose.Cells for .NET. | Add a column chart only after confirming that every cell in the source range contains data, then set the series and save the workbook with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# check for blank cells in a range before adding a chart | prevent chart generation when Excel source range contains empty values using Aspose.Cells | validate data source range for null values before creating a chart in Aspose.Cells .NET | how to stop chart creation if any cell in the source range is empty in C#
// Tags: Aspose.Cells range empty cell check | C# column chart creation after validation | Excel data source integrity Aspose.Cells | Aspose.Cells chart generation with non‑blank range | validate worksheet range before chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// The example loads or creates an Excel workbook, defines a data range (A1:B5), scans the cells for null values, aborts chart creation if a blank is found, otherwise adds a column chart linked to the validated range, sets a title, and saves the workbook.
class ChartGenerator
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Ensure the input workbook exists; create a simple one if missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Populate sample data in A1:B5
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("C");
                ws.Cells["B4"].PutValue(30);
                ws.Cells["A5"].PutValue("D");
                ws.Cells["B5"].PutValue(40);
                workbook.Save(inputPath);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the chart (e.g., A1:B5)
            const string dataRangeAddress = "A1:B5";
            AsposeRange dataRange = sheet.Cells.CreateRange(dataRangeAddress);

            // Validate that the range does not contain empty cells
            bool hasEmptyCell = false;
            foreach (Cell cell in dataRange)
            {
                // In Aspose.Cells, a cell is considered empty if its Value is null
                if (cell.Value == null)
                {
                    hasEmptyCell = true;
                    Console.WriteLine($"Empty cell found at {cell.Name}");
                    break;
                }
            }

            if (hasEmptyCell)
            {
                Console.WriteLine("Chart generation aborted due to empty cells in the data range.");
            }
            else
            {
                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart's data source to the validated range
                chart.NSeries.Add(dataRangeAddress, true);

                // Optional: set chart title
                chart.Title.Text = "Sample Chart";

                // Save the workbook with the new chart
                workbook.Save(outputPath);
                Console.WriteLine("Chart created and workbook saved successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
