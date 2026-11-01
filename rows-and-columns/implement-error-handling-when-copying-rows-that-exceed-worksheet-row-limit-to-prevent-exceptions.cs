// Title: C# Aspose.Cells example: copy rows between worksheets with worksheet row‑limit validation and robust error handling
// AI Prompts: Create C# code that uses Aspose.Cells to copy a specified range of rows from one worksheet to another, automatically truncating the range if the destination start row plus the range exceeds the worksheet's MaxRow, and includes checks for file existence and worksheet count. | Write a reusable C# method that validates the destination start row against Excel's maximum row count, adjusts the rowsToCopy value accordingly, performs the copy with Aspose.Cells, and logs meaningful error messages for missing files, insufficient worksheets, or overflow conditions.
// Common Searches: Aspose.Cells copy rows with MaxRow check in C# | prevent Excel row overflow when copying rows using Aspose.Cells .NET | adjust rowsToCopy based on destination start row Aspose.Cells | C# error handling for large row copy with Aspose.Cells | how to validate worksheet row limit before copying rows Aspose.Cells
// Tags: Aspose.Cells row copy MaxRow validation | C# worksheet row overflow protection | Aspose.Cells adjust rowsToCopy for Excel limit | C# error handling Aspose.Cells copy operation | Aspose.Cells copy rows between worksheets

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates copying rows from one worksheet to another with Aspose.Cells in C#, checking the destination worksheet's MaxRow, automatically adjusting the number of rows to copy to stay within Excel's row limit, and handling errors such as missing files, insufficient worksheets, and overflow conditions.
class RowCopyWithLimitCheck
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Ensure there are at least two worksheets
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook must contain at least two worksheets.");
                return;
            }

            // Get source and destination worksheets
            Worksheet srcSheet = workbook.Worksheets[0];
            Worksheet destSheet = workbook.Worksheets[1];

            // Define copy parameters (zero‑based indices)
            int srcStartRow = 0;          // first row to copy from source
            int rowsToCopy = 500000;      // intended number of rows to copy
            int destStartRow = 600000;    // first row in destination where data will be pasted

            // Maximum rows allowed in the destination worksheet (total count, not zero‑based index)
            int maxRows = destSheet.Cells.MaxRow + 1; // typically 1,048,576 for .xlsx

            // Calculate the last row index after copy
            long lastDestRow = (long)destStartRow + rowsToCopy - 1;

            // Adjust rowsToCopy if the operation would exceed the worksheet limit
            if (lastDestRow >= maxRows)
            {
                rowsToCopy = (int)(maxRows - destStartRow);
                if (rowsToCopy <= 0)
                {
                    Console.WriteLine("Copy operation aborted: destination start row exceeds worksheet row limit.");
                    return;
                }

                Console.WriteLine($"Adjusted rows to copy to {rowsToCopy} to stay within the worksheet limit.");
            }

            // Determine the number of columns to copy (all columns that contain data in the source sheet)
            int totalColumns = srcSheet.Cells.MaxColumn + 1;

            // Create source and destination ranges
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange(srcStartRow, 0, rowsToCopy, totalColumns);
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange(destStartRow, 0, rowsToCopy, totalColumns);

            // Perform the copy
            destRange.Copy(srcRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Rows copied successfully. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
