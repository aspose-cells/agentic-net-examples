// Title: How to check for empty cells in a worksheet range before generating a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Iterate over a specified CellArea in a workbook and abort chart creation if any cell is null or empty, using the Aspose.Cells C# API. | Create a column chart only after confirming that every cell in the source range contains a value, then save the workbook with Aspose.Cells. | Add pre‑chart validation that logs the address of the first blank cell and prevents adding NSeries to the chart in C#.
// Common Searches: C# Aspose.Cells verify that a range has no blank cells before adding a chart | How to skip chart generation in Aspose.Cells when source data contains empty values | Check Excel range for null or empty strings using Aspose.Cells API | Prevent Aspose.Cells column chart from using incomplete data in .NET
// Tags: Aspose.Cells validate chart data range | C# check empty cells before chart creation | Aspose.Cells column chart from non‑empty range | Excel workbook range integrity Aspose.Cells | pre‑chart validation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, defines a cell range (A1:B5), scans the range for null or empty cells, aborts chart creation if a blank is found (logging its address), otherwise adds a column chart using the validated range, sets a title, and saves the workbook as Output.xlsx.
class ChartGenerator
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the chart (A1:B5)
            CellArea dataRange = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 4,
                EndColumn = 1
            };

            // Validate that the range does not contain empty cells
            bool hasEmpty = false;
            for (int row = dataRange.StartRow; row <= dataRange.EndRow && !hasEmpty; row++)
            {
                for (int col = dataRange.StartColumn; col <= dataRange.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                    {
                        hasEmpty = true;
                        Console.WriteLine($"Empty cell found at {cell.Name}.");
                        break;
                    }
                }
            }

            if (hasEmpty)
            {
                Console.WriteLine("Chart generation aborted due to empty cells in the data range.");
                return;
            }

            // Create a column chart positioned from row 5, column 0 to row 20, column 5
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Build the address string for the range (e.g., "A1:B5")
            string startAddress = CellsHelper.CellIndexToName(dataRange.StartRow, dataRange.StartColumn);
            string endAddress = CellsHelper.CellIndexToName(dataRange.EndRow, dataRange.EndColumn);
            string rangeAddress = $"{startAddress}:{endAddress}";

            // Set the chart's data source to the validated range
            chart.NSeries.Add(rangeAddress, true);

            // Optional: set chart title
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook with the new chart
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Chart created and workbook saved to {outputPath}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
