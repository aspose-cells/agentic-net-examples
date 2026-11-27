// Title: How to validate an Excel sparkline data range for error cells before creating the sparkline using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# check for error values in a sparkline source range | prevent sparkline creation when source range contains #DIV/0! using Aspose.Cells | detect CellValueType.IsError in Excel range before adding sparkline with Aspose.Cells .NET | validate sparkline data range for formula errors in C# Aspose.Cells example | skip sparkline group creation if source cells have errors Aspose.Cells
// Tags: Aspose.Cells sparkline range validation | C# detect error cells Aspose.Cells | Excel sparkline creation error handling .NET | CellValueType.IsError usage Aspose.Cells | prevent sparkline on error cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineValidationDemo
{
    // The example creates a workbook, inserts numeric data and a formula that generates a #DIV/0! error, then scans the defined range (A1:D2) for error cells using CellValueType.IsError. If no errors are found, it adds a line‑type sparkline group to column E, customizes high/low points, and saves the file as SparklineValidationResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Initialize a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate test data (including an intentional error)
                // Row 0: valid numeric data
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["C1"].PutValue(30);
                cells["D1"].PutValue(40);

                // Row 1: introduce a formula error (#DIV/0!)
                cells["A2"].Formula = "=1/0"; // This will produce a division by zero error
                cells["B2"].PutValue(5);
                cells["C2"].PutValue(15);
                cells["D2"].PutValue(25);

                // Define the data range for the sparkline (including the error row)
                string dataRange = "A1:D2";

                // Validate the data range for errors before creating the sparkline
                if (ContainsError(cells, dataRange))
                {
                    Console.WriteLine($"The data range \"{dataRange}\" contains error cells. Sparkline will not be created.");
                }
                else
                {
                    // Define where the sparkline will be placed (column E, rows 1 and 2)
                    CellArea sparklineArea = new CellArea
                    {
                        StartRow = 0,
                        EndRow = 1,
                        StartColumn = 4,
                        EndColumn = 4
                    };

                    // Add a sparkline group (Line type) using the validated range
                    int groupIndex = sheet.SparklineGroups.Add(
                        SparklineType.Line,   // Sparkline type
                        dataRange,            // Data range (already validated)
                        false,                // Do not plot right‑to‑left
                        sparklineArea);       // Destination cells for the sparklines

                    SparklineGroup group = sheet.SparklineGroups[groupIndex];

                    // Optional: customize the sparkline appearance
                    group.ShowHighPoint = true;
                    group.ShowLowPoint = true;
                    group.HighPointColor.Color = System.Drawing.Color.Green;
                    group.LowPointColor.Color = System.Drawing.Color.Red;
                    group.LineWeight = 1.0;

                    Console.WriteLine("Sparkline group created successfully.");
                }

                // Save the workbook
                string outputPath = "SparklineValidationResult.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="cells">The Cells collection of the worksheet.</param>
        /// <param name="range">A string representing the range (e.g., "A1:D5").</param>
        /// <returns>True if an error is found; otherwise, false.</returns>
        private static bool ContainsError(Cells cells, string range)
        {
            // Split the range into start and end cells and create a CellArea
            string[] parts = range.Split(':');
            if (parts.Length != 2)
                throw new ArgumentException("Invalid range format.", nameof(range));

            CellArea area = CellArea.CreateCellArea(parts[0], parts[1]);

            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    Cell cell = cells[row, col];
                    // Cell.Type returns a CellValueType enum; IsError indicates an error cell
                    if (cell.Type == CellValueType.IsError)
                    {
                        Console.WriteLine($"Error found at {cell.Name} with value \"{cell.StringValue}\".");
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
