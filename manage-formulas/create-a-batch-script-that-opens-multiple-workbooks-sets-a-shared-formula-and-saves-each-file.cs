// Title: C# batch script to open every .xlsx workbook in a folder, apply a shared formula to column C rows 1‑10, and save the files with Aspose.Cells
// AI Prompts: Write a C# console program that enumerates all .xlsx files in a specified directory, loads each workbook with Aspose.Cells, and inserts a formula `A{row}+B{row}` into cells C1:C10 before saving the workbook in place. | Refactor the existing per‑cell loop into a single Aspose.Cells Range that uses one formula for the entire C1:C10 block, while still processing every workbook in the folder. | Enhance the script to log detailed error information for files that cannot be opened, skip those files, and continue processing the remaining workbooks without terminating the application.
// Common Searches: how to set a shared formula across multiple rows using Aspose.Cells in C# | batch modify column C formula in all Excel files in a folder with Aspose.Cells .NET | C# script to apply A+B formula to C1:C10 for each workbook in a directory | process many .xlsx files and overwrite them after adding formulas using Aspose.Cells | error handling for file‑wise workbook processing Aspose.Cells C#
// Tags: batch apply formula range Aspose.Cells | process multiple .xlsx files C# | set column C formula rows 1-10 Aspose.Cells | overwrite workbook after formula insertion | file‑wise error handling Aspose.Cells | range based formula assignment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Scans a folder for .xlsx workbooks, loads each with Aspose.Cells, assigns a formula that adds column A and B into cells C1‑C10, saves the workbook back to its original location, and logs any file‑specific errors while continuing processing.
class Program
{
    static void Main()
    {
        try
        {
            // Folder that contains the workbooks to process
            string inputFolder = @"C:\InputFolder";

            // Verify that the input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Retrieve all .xlsx files in the specified folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

            foreach (string filePath in workbookFiles)
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook from disk
                    Workbook workbook = new Workbook(filePath);

                    // Work with the first worksheet (index 0)
                    Worksheet sheet = workbook.Worksheets[0];

                    // Apply formula to each row in column C (C1:C10)
                    for (int row = 0; row < 10; row++)
                    {
                        // Build a relative formula for the current row (e.g., A1+B1, A2+B2, ...)
                        string formula = $"A{row + 1}+B{row + 1}";
                        // Set the formula for the cell in column C (index 2)
                        sheet.Cells[row, 2].Formula = formula;
                    }

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath);
                    Console.WriteLine($"Processed and saved: {filePath}");
                }
                catch (Exception exFile)
                {
                    // Log errors related to a specific file but continue processing others
                    Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
