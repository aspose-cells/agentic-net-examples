// Title: C# example: split column C into three 5‑character fixed‑width fields with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split Excel column into fixed width parts | How to use TextToColumns for fixed width on column C in Aspose.Cells | Divide cell text into equal 5‑character segments using Aspose.Cells .NET | C# code to extract substrings from column C and write to adjacent columns in Excel | Aspose.Cells .NET example for fixed‑length field extraction from a worksheet
// Tags: Aspose.Cells fixed-width TextToColumns C# | extract fixed-length substrings Excel Aspose.Cells | write split parts to adjacent columns Aspose.Cells | C# Excel column segmentation Aspose.Cells | process column C fixed-width fields .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, finds the data range in column C, and for each cell extracts three substrings of five characters each. The parts are written back to columns C, D, and E, and the workbook is saved as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple one if missing.
                if (!File.Exists(inputPath))
                {
                    CreateSampleWorkbook(inputPath);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the data range in column C (zero‑based index 2).
                int startRow = sheet.Cells.GetFirstDataRow(2);
                int endRow = sheet.Cells.GetLastDataRow(2);

                // If no data found, exit gracefully.
                if (startRow < 0 || endRow < 0)
                {
                    Console.WriteLine("No data found in column C.");
                    return;
                }

                // Split each cell into three fixed‑width parts (5 characters each).
                const int fieldCount = 3;
                const int fieldWidth = 5;

                for (int row = startRow; row <= endRow; row++)
                {
                    Cell sourceCell = sheet.Cells[row, 2]; // Column C
                    string cellText = sourceCell.StringValue ?? string.Empty;

                    for (int i = 0; i < fieldCount; i++)
                    {
                        int startIdx = i * fieldWidth;
                        string part = startIdx < cellText.Length
                            ? cellText.Substring(startIdx, Math.Min(fieldWidth, cellText.Length - startIdx))
                            : string.Empty;

                        // Write the part to the adjacent columns (C, D, E).
                        sheet.Cells[row, 2 + i].PutValue(part);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a simple workbook with sample data in column C for demonstration.
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Sample strings longer than 15 characters to demonstrate splitting.
            ws.Cells[0, 2].PutValue("ABCDEFGHIJKLMNO");
            ws.Cells[1, 2].PutValue("123456789012345");
            ws.Cells[2, 2].PutValue("HelloWorldAgain!");

            wb.Save(path);
        }
    }
}
