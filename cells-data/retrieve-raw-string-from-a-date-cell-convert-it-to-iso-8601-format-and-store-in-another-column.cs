// Title: Read a date from an Excel cell, convert it to ISO 8601 format, and write the string to another column using Aspose.Cells for .NET
// AI Prompts: Create C# code with Aspose.Cells that extracts a DateTime from cell A1, formats it as an ISO 8601 string, and places the result in cell B1. | Modify the example to loop through all rows in column A, convert each valid date to ISO 8601, and write the output to the same row in column B. | Add error handling that logs a warning and skips rows where the source cell does not contain a parsable date.
// Common Searches: Aspose.Cells C# convert Excel date cell to ISO 8601 string | How to read a date from an Excel cell and write formatted ISO 8601 string to another column using Aspose.Cells | Parse cell value as DateTime and output ISO 8601 format with Aspose.Cells .NET | Save ISO 8601 date string in Excel workbook using Aspose.Cells for .NET
// Tags: convert Excel date cell to ISO 8601 Aspose.Cells | read DateTime cell C# Aspose.Cells | write ISO 8601 string to column B Aspose.Cells | skip invalid date rows Aspose.Cells | round‑trip ISO 8601 format .NET Excel

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDateConversion
{
    // Loads input.xlsx (creates a sample if missing), reads the date from cell A1, formats it as an ISO 8601 string, writes the string to cell B1, and saves the result as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a sample if missing
                if (!File.Exists(inputPath))
                {
                    var sampleWb = new Workbook();
                    sampleWb.Worksheets[0].Cells["A1"].PutValue(DateTime.Now);
                    sampleWb.Save(inputPath);
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];
                var sourceCell = worksheet.Cells["A1"];

                DateTime dateValue;

                // Determine if the cell actually contains a DateTime value
                if (sourceCell.Type == CellValueType.IsDateTime)
                {
                    dateValue = sourceCell.DateTimeValue;
                }
                else
                {
                    // Attempt to parse the cell's string representation
                    if (!DateTime.TryParse(sourceCell.StringValue, out dateValue))
                    {
                        Console.WriteLine("Cell A1 does not contain a valid DateTime value.");
                        return;
                    }
                }

                // Convert to ISO 8601 (round‑trip) format
                string iso8601String = dateValue.ToString("o");

                // Store the ISO string in B1
                worksheet.Cells["B1"].PutValue(iso8601String);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Conversion completed. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
