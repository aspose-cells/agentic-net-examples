// Title: Validate empty destination cells and move range H1:H5 to I1:I5 using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that checks whether cells I1:I5 are empty before copying the values from H1:H5 and then clears the original range. | Create a reusable C# method that detects existing data in a target range, copies a source range to it, and removes the source cells using Aspose.Cells.
// Common Searches: Aspose.Cells how to ensure destination range is empty before moving cells in C# | C# copy Excel range H1:H5 to I1:I5 without overwriting existing data using Aspose.Cells | detect overlapping data when moving a column range with Aspose.Cells .NET | validate empty cells in Excel before transferring a range with Aspose.Cells library | move column data to adjacent column safely Aspose.Cells C#
// Tags: Aspose.Cells validate empty destination range | move range without overlap Aspose.Cells | copy range and clear source Aspose.Cells | check destination cells empty C# | range overlap detection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range (C# 8+)
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, checks that the target cells I1:I5 contain no data, copies the values from H1:H5 to I1:I5 with Aspose.Cells, clears the original H1:H5 range, and saves the file, reporting an error if the destination is not empty.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Source range H1:H5 (zero‑based indices)
            int srcStartRow = 0;      // Row 1
            int srcStartColumn = 7;   // Column H
            int srcRowCount = 5;      // Rows 1‑5
            int srcColumnCount = 1;   // Single column

            // Destination range I1:I5
            int destStartRow = 0;     // Row 1
            int destStartColumn = 8;  // Column I

            // Ensure destination cells are empty
            bool overlap = false;
            for (int r = 0; r < srcRowCount; r++)
            {
                Cell destCell = sheet.Cells[destStartRow + r, destStartColumn];
                if (destCell.Value != null && !string.IsNullOrEmpty(destCell.StringValue))
                {
                    overlap = true;
                    break;
                }
            }

            if (overlap)
            {
                Console.WriteLine("Cannot move range H1:H5 to I1:I5 because destination overlaps existing data.");
            }
            else
            {
                // Copy source range to destination
                AsposeRange srcRange = sheet.Cells.CreateRange(srcStartRow, srcStartColumn, srcRowCount, srcColumnCount);
                AsposeRange destRange = sheet.Cells.CreateRange(destStartRow, destStartColumn, srcRowCount, srcColumnCount);
                destRange.Copy(srcRange);

                // Clear the original source range
                sheet.Cells.ClearRange(srcStartRow, srcStartColumn, srcRowCount, srcColumnCount);

                Console.WriteLine("Range moved successfully.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
