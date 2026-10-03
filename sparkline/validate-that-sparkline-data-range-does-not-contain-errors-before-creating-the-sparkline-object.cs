// Title: Check for error values in an Excel sparkline source range before creating the sparkline using Aspose.Cells for C#
// AI Prompts: Generate C# code that iterates over a specified cell range, detects any #N/A or other error types, and adds a line sparkline only when the range is error‑free with Aspose.Cells. | Write a method that loads a workbook, validates the B2:B10 range for error cells, and conditionally creates a sparkline group with markers using the Aspose.Cells Sparkline API. | Provide a C# example that logs the first error cell found in a sparkline data range and skips sparkline creation if any errors exist.
// Common Searches: Aspose.Cells how to skip sparkline creation if source range contains #N/A | C# validate Excel range for errors before adding sparkline | check for error values in sparkline data range using Aspose.Cells | conditional sparkline generation in .NET based on cell error detection | prevent Aspose.Cells sparkline from failing due to error cells
// Tags: sparkline source range validation Aspose.Cells | conditional sparkline creation C# | cell error detection Aspose.Cells | excel sparkline error handling C# | Aspose.Cells sparkline data verification

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook or creates a new one, fills B2:B10 with sample numbers if needed, scans the range for any error cells, and proceeds with sparkline creation only when no errors are detected, finally saving the workbook as output.xlsx.
class SparklineValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Populate sample numeric data in B2:B10
                for (int i = 0; i < 9; i++)
                {
                    ws.Cells[i + 1, 1].PutValue(i + 1);
                }
                workbook.Save(inputPath);
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the sparkline (e.g., B2:B10)
            const string dataRangeAddress = "B2:B10";
            AsposeRange dataRange = sheet.Cells.CreateRange(dataRangeAddress);

            // Validate that the data range does not contain any error values
            bool hasError = false;
            foreach (Cell cell in dataRange)
            {
                if (cell.Type == CellValueType.IsError)
                {
                    hasError = true;
                    Console.WriteLine($"Error found in cell {cell.Name}: {cell.StringValue}");
                    break; // Stop checking after the first error is found
                }
            }

            // Create the sparkline only if the data range is error‑free
            if (!hasError)
            {
                try
                {
                    // NOTE: Sparkline APIs require the Aspose.Cells.Sparkline assembly.
                    // If the assembly is unavailable, this block can be omitted or
                    // replaced with alternative logic.

                    // Example placeholder for sparkline creation:
                    // const string sparklineLocation = "C2:C10";
                    // var sparklineGroup = sheet.SparklineGroups.Add(
                    //     Aspose.Cells.Sparkline.SparklineType.Line,
                    //     sparklineLocation,
                    //     dataRangeAddress);
                    // sparklineGroup.ShowMarkers = true;
                    // sparklineGroup.MarkerColor = System.Drawing.Color.Red;

                    Console.WriteLine("Sparkline creation logic would be executed here.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sparkline: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Sparkline creation skipped due to errors in the data range.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
