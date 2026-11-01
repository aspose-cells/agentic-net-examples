// Title: Copy rows from one worksheet to another and auto‑fit rows while preserving original row heights using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy rows from source worksheet to destination worksheet preserving row height | how to auto‑fit rows after copying them with Aspose.Cells in a .NET project | copy a range of rows between worksheets and keep original row heights using Aspose.Cells for .NET | set row height programmatically after copying rows with Aspose.Cells C# example | Aspose.Cells copy rows and auto‑fit rows programmatically
// Tags: copy rows between worksheets Aspose.Cells C# | preserve row height Aspose.Cells API | auto‑fit rows Aspose.Cells method | copy cell values and styles Aspose.Cells | set row height programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRowCopyExample
{
    // The example loads an existing workbook, copies a specified block of rows from a "Source" worksheet to a "Destination" worksheet by iterating each cell, replicates values, styles, and row heights, then auto‑fits the pasted rows and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get source and destination worksheets (by name; adjust if needed)
                Worksheet sourceSheet = workbook.Worksheets["Source"];
                Worksheet destSheet = workbook.Worksheets["Destination"];

                if (sourceSheet == null || destSheet == null)
                {
                    Console.WriteLine("Error: Specified worksheets not found in the workbook.");
                    return;
                }

                // Define the range of rows to copy (zero‑based indexes)
                int sourceStartRow = 0;   // first row to copy from source
                int destStartRow   = 5;   // first row to paste into destination
                int rowCount       = 10;  // number of rows to copy

                // Copy rows one by one
                for (int i = 0; i < rowCount; i++)
                {
                    int srcRowIndex = sourceStartRow + i;
                    int dstRowIndex = destStartRow   + i;

                    // Copy each cell in the current row
                    for (int col = 0; col < sourceSheet.Cells.MaxColumn; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[srcRowIndex, col];
                        Cell dstCell = destSheet.Cells[dstRowIndex, col];

                        // Copy value
                        dstCell.PutValue(srcCell.Value);

                        // Copy style (including font, alignment, borders, etc.)
                        dstCell.SetStyle(srcCell.GetStyle());
                    }

                    // Copy the row height from source to destination
                    destSheet.Cells.Rows[dstRowIndex].Height = sourceSheet.Cells.Rows[srcRowIndex].Height;
                }

                // Optional: auto‑fit the destination rows (keeps the height set above if content fits)
                destSheet.AutoFitRows(destStartRow, destStartRow + rowCount - 1);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
