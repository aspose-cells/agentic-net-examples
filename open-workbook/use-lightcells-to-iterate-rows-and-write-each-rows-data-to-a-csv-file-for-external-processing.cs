// Title: Iterate rows of an XLSX workbook with Aspose.Cells and export to CSV in C#
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, uses the LightCells API to walk through each row of the used range, and writes the cell values to a CSV stream applying RFC 4180 escaping. | Create a method that receives a worksheet, determines its MaxDataRow and MaxDataColumn, and outputs a CSV line for every row while handling commas, quotes, and line breaks correctly.
// Common Searches: Aspose.Cells C# export large Excel sheet to CSV using row-by-row processing | How to write CSV from an Excel workbook with proper quoting in .NET Aspose.Cells | LightCells API example for streaming XLSX data to CSV in C# | C# convert XLSX to CSV without loading the entire worksheet into memory Aspose.Cells | Iterate worksheet rows and generate CSV file with Aspose.Cells LoadOptions
// Tags: lightcells row iteration csv export | aspose.cells export worksheet to csv | c# csv escaping for excel values | stream xlsx rows to csv .net | rfc4180 compliant csv generation aspnet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, determines the used range, iterates each cell row‑by‑row using LightCells, applies CSV escaping for commas, quotes and newlines, and writes the formatted rows to an output CSV file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            using (var workbook = new Workbook(inputPath, loadOptions))
            {
                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Determine the used range
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                // Open CSV writer
                using (var csvWriter = new StreamWriter(outputPath))
                {
                    // Iterate through each row in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        var rowValues = new List<string>();

                        // Iterate through each column in the current row
                        for (int col = 0; col <= maxCol; col++)
                        {
                            // Retrieve the cell value as a string (handles nulls)
                            string cellText = worksheet.Cells[row, col].StringValue ?? string.Empty;

                            // Escape double quotes by doubling them
                            if (cellText.Contains("\""))
                                cellText = cellText.Replace("\"", "\"\"");

                            // Enclose the value in quotes if it contains commas, quotes, or newlines
                            if (cellText.Contains(",") || cellText.Contains("\r") || cellText.Contains("\n") || cellText.Contains("\""))
                                cellText = $"\"{cellText}\"";

                            rowValues.Add(cellText);
                        }

                        // Write the CSV line
                        csvWriter.WriteLine(string.Join(",", rowValues));
                    }
                }
            }

            Console.WriteLine("CSV export completed successfully.");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
