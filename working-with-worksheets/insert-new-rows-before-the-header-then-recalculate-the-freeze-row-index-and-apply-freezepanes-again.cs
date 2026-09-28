// Title: Insert rows before a header row and adjust FreezePanes in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a specified number of rows above a given header row, recalculate the freeze‑pane row index, and apply FreezePanes with the 4‑parameter overload in C# using Aspose.Cells. | Load an existing workbook, verify the input file, add rows before the header, shift the frozen rows to stay above the new header, ensure the output directory exists, and save the updated file.
// Common Searches: C# Aspose.Cells insert rows before header while preserving frozen panes | How to recalculate FreezePanes row index after adding rows in Aspose.Cells .NET | Aspose.Cells example for inserting rows above a header and reapplying FreezePanes | Adjust frozen rows after inserting rows with Aspose.Cells in C# | Insert multiple rows before a specific row and maintain freeze pane using Aspose.Cells
// Tags: insert rows before header Aspose.Cells | recalculate freeze pane row index Aspose.Cells | apply FreezePanes 4‑parameter overload Aspose.Cells | load and save workbook Aspose.Cells C# | verify input file existence C#

using System;
using System.IO;
using Aspose.Cells;

// The code loads input.xlsx, inserts two rows before the original header row (row 6), updates the FreezePanes to the new header position using the 4‑parameter overload, creates the output directory if needed, and saves the modified workbook as output.xlsx while handling missing input files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets.Count > 0 ? workbook.Worksheets[0] : null;
            if (sheet == null)
            {
                Console.WriteLine("No worksheets found in the workbook.");
                return;
            }

            // Index of the header row (0‑based). Adjust as needed.
            int headerRowIndex = 5; // e.g., Excel row 6

            // Number of rows to insert before the header
            int rowsToInsert = 2;

            // Insert the rows
            sheet.Cells.InsertRows(headerRowIndex, rowsToInsert);

            // Recalculate the freeze pane row index.
            // If the original freeze was at the header row, shift it down by the inserted rows.
            int originalFreezeRow = headerRowIndex;          // original freeze row before insertion
            int newFreezeRow = originalFreezeRow + rowsToInsert;

            // Apply FreezePanes again (freeze rows above newFreezeRow, no column freeze)
            // Use the 4‑parameter overload: FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(newFreezeRow, 0, newFreezeRow, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
