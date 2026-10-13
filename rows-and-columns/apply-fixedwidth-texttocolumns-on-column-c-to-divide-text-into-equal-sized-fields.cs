// Title: Split column C into three 10‑character fields using Aspose.Cells TextToColumns in C#
// AI Prompts: Write C# code that loads an Excel workbook, reads each cell in column C, and applies a fixed‑width TextToColumns operation to divide the text into three 10‑character columns starting at column D with Aspose.Cells. | Generate a C# example that iterates over all rows, uses widths {10,10,10} on column C, trims each extracted segment, and saves the updated workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to apply fixed width TextToColumns to a specific column | split Excel column into multiple columns by character count using Aspose.Cells | C# code sample for dividing text in column C into three equal parts with Aspose.Cells | extract substrings of length 10 from column C and write to columns D E F in .xlsx using Aspose.Cells | process all rows in a worksheet with MaxDataRow and fixed-width columns Aspose.Cells
// Tags: fixed-width column splitting Aspose.Cells C# | split column C into three columns Aspose.Cells | extract equal-length substrings Excel C# | iterate rows MaxDataRow Aspose.Cells | write trimmed parts to adjacent cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, reads each cell in column C, splits its text into three 10‑character segments, trims each segment, writes the results into columns D, E, and F, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int lastRow = worksheet.Cells.MaxDataRow;

            // Fixed width segments (10 characters each)
            int[] widths = new int[] { 10, 10, 10 };
            int startColumn = 2; // Column C (zero‑based index)
            int destinationStartColumn = 3; // Column D

            // Process each row in the source range
            for (int row = 0; row <= lastRow; row++)
            {
                string sourceText = worksheet.Cells[row, startColumn].StringValue ?? string.Empty;

                for (int i = 0; i < widths.Length; i++)
                {
                    int startPos = i * widths[i];
                    string part = sourceText.Length > startPos
                        ? sourceText.Substring(startPos, Math.Min(widths[i], sourceText.Length - startPos))
                        : string.Empty;

                    // Trim the part and place it into the destination cell
                    worksheet.Cells[row, destinationStartColumn + i].PutValue(part.Trim());
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
