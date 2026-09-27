// Title: Hide columns by header name and export the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing XLSX file, scans the first row for specified header strings (case‑insensitive), hides the matching columns, and saves the result to a new file. | Create a helper method in C# that accepts an input workbook path, a list of header names to conceal, and an output path, then performs column hiding based on those headers using Aspose.Cells and writes the modified workbook.
// Common Searches: Aspose.Cells hide columns based on header values in C# | C# hide multiple worksheet columns using a header list with Aspose.Cells | Save a modified Excel file to a different location using Aspose.Cells .NET | Case‑insensitive header matching for column visibility in Aspose.Cells
// Tags: column hiding by header Aspose.Cells | write updated workbook Aspose.Cells | case‑insensitive header matching Aspose.Cells | ensure output folder exists Aspose.Cells | column visibility API .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Loads an XLSX workbook, iterates the first row to find headers listed in an array, hides those columns (case‑insensitive), creates the output directory if missing, and saves the modified workbook to a new XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = @"C:\Input\Sample.xlsx";
            const string outputPath = @"C:\Output\ModifiedSample.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Headers whose columns should be hidden
                string[] headersToHide = { "Secret", "Internal", "Confidential" };

                // Scan the first row for matching headers
                Cells cells = sheet.Cells;
                int maxColumn = cells.MaxColumn; // last used column index

                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell headerCell = cells[0, col];
                    if (headerCell.Type == CellValueType.IsString)
                    {
                        string headerText = headerCell.StringValue.Trim();

                        foreach (string hideHeader in headersToHide)
                        {
                            if (string.Equals(headerText, hideHeader, StringComparison.OrdinalIgnoreCase))
                            {
                                // Hide the entire column
                                sheet.Cells.HideColumn(col);
                                break;
                            }
                        }
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
