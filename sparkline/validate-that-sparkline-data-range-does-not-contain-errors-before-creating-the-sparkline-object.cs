// Title: How to validate a sparkline source range for error values before creating the sparkline with Aspose.Cells in C#
// AI Prompts: Iterate over a worksheet range, detect cells where CellValueType.IsError is true, and add a SparklineGroup only when the range is error‑free using Aspose.Cells for .NET. | Create a line sparkline at a specified cell after confirming the source range contains no #DIV/0! or #N/A errors, then save the workbook.
// Common Searches: Aspose.Cells C# check for error cells before adding a sparkline | prevent sparkline creation if source range contains #DIV/0! in .NET | validate Excel range for errors using Aspose.Cells before sparkline group addition | C# example of iterating over a range to detect IsError values with Aspose.Cells | how to abort sparkline generation when data range has errors in Aspose.Cells
// Tags: sparkline source range error validation C# | Aspose.Cells check cell IsError before sparkline | create sparkline after range verification Aspose.Cells | line sparkline group addition conditional C# | Excel workbook save after sparkline validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSparklineValidation
{
    // The example creates a workbook, fills cells A1:A5 with numeric data, scans the range for any error‑type cells, aborts sparkline creation if an error is found, otherwise adds a line sparkline at B1, reports the outcome, saves the file as SparklineValidated.xlsx, and handles exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(7);
                sheet.Cells["A4"].PutValue(2);
                sheet.Cells["A5"].PutValue(9);

                // Define the data range that will be used for the sparkline
                string dataRange = "A1:A5";

                // Validate that the data range does not contain any error values
                bool hasError = false;
                AsposeRange range = sheet.Cells.CreateRange(dataRange);
                foreach (Cell cell in range)
                {
                    // Check if the cell contains an error (e.g., #DIV/0!, #N/A, etc.)
                    if (cell.Type == CellValueType.IsError)
                    {
                        hasError = true;
                        Console.WriteLine($"Error found in cell {cell.Name}: {cell.StringValue}");
                        break;
                    }
                }

                if (hasError)
                {
                    Console.WriteLine("Sparkline creation aborted due to errors in the data range.");
                }
                else
                {
                    // Define where the sparkline will be placed
                    CellArea location = new CellArea
                    {
                        StartRow = 0,
                        EndRow = 0,
                        StartColumn = 1,
                        EndColumn = 1
                    };

                    // Add the sparkline group
                    int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
                    SparklineGroup group = sheet.SparklineGroups[groupIdx];

                    // Access the created sparkline and display its data range
                    Sparkline sparkline = group.Sparklines[0];
                    Console.WriteLine("Sparkline created successfully. DataRange: " + sparkline.DataRange);
                }

                // Save the workbook (lifecycle rule)
                string outputPath = "SparklineValidated.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
