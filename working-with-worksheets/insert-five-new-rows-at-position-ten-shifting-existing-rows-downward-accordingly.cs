// Title: Insert five rows at row 10 in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert multiple rows at a specific index and shift down | C# code to add five rows starting at row 10 in an existing .xlsx file using Aspose.Cells | How to use Worksheet.Cells.InsertRows to insert rows in a .NET Excel workbook | Save modified workbook after inserting rows with Aspose.Cells in C# | Insert rows at zero‑based index 9 with Aspose.Cells and write values to new rows
// Tags: Worksheet.Cells.InsertRows method C# | add five rows to Excel worksheet Aspose.Cells | zero‑based row indexing in Aspose.Cells | populate newly inserted rows with values C# | save modified .xlsx workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace InsertRowsExample
{
    // The example loads a source .xlsx file using Aspose.Cells, accesses the first worksheet, inserts five rows at the human‑readable row 10 (zero‑based index 9) which shifts existing rows downward, writes a label into each newly created row, and saves the updated workbook to the specified destination path.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source Excel file (required)
        /// args[1] - path to the destination Excel file (optional, defaults to "output.xlsx" in the same folder)
        /// </param>
        static void Main(string[] args)
        {
            try
            {
                // Validate input arguments
                if (args == null || args.Length == 0)
                {
                    Console.WriteLine("Usage: InsertRowsExample <sourceFilePath> [destinationFilePath]");
                    return;
                }

                string sourcePath = args[0];
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Source file not found at \"{sourcePath}\"");
                    return;
                }

                // Determine destination path
                string destPath;
                if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
                {
                    destPath = args[1];
                }
                else
                {
                    string directory = Path.GetDirectoryName(sourcePath) ?? Directory.GetCurrentDirectory();
                    destPath = Path.Combine(directory, "output.xlsx");
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                // Access the first worksheet (you can modify this to target a specific sheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Insert five rows at position ten.
                // Aspose.Cells uses zero‑based indexing, so row 10 corresponds to index 9.
                int insertAtRowIndex = 9; // Row 10 (human readable)
                int numberOfRowsToInsert = 5;

                // Perform the insertion. Existing rows from the insert index downward are shifted.
                worksheet.Cells.InsertRows(insertAtRowIndex, numberOfRowsToInsert);

                // Optional: Write a message into the newly inserted rows to demonstrate they exist.
                for (int i = 0; i < numberOfRowsToInsert; i++)
                {
                    int row = insertAtRowIndex + i;
                    worksheet.Cells[row, 0].PutValue($"Inserted Row {row + 1}");
                }

                // Save the modified workbook
                workbook.Save(destPath, SaveFormat.Xlsx);

                Console.WriteLine($"Successfully inserted {numberOfRowsToInsert} rows at position 10.");
                Console.WriteLine($"Modified workbook saved to \"{destPath}\"");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
