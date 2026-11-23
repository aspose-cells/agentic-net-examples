// Title: Batch apply a shared SUM formula to all .xlsx workbooks in a folder using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set shared formula for a range across multiple workbooks | process every .xlsx file in a directory and add a shared formula with Aspose.Cells | apply SUM(B1:B10) as a shared formula to column A in many Excel files programmatically | iterate through Excel workbooks folder and set shared formulas using Aspose.Cells C# | batch update Excel workbooks to use shared formulas in Aspose.Cells .NET
// Tags: batch shared formula Aspose.Cells C# | apply shared formula to range Aspose.Cells | process multiple .xlsx workbooks C# | set shared formula programmatically Aspose.Cells | update Excel files in folder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, sets the formula SUM(B1:B10) as a shared formula in cells A1:A10 of the first worksheet, saves the workbook back to its original location, and logs the processing outcome while handling missing files and exceptions.
class BatchSharedFormula
{
    static void Main()
    {
        try
        {
            // Folder containing the workbooks
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files in the folder (you can adjust the pattern as needed)
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");

            foreach (string filePath in files)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define the range where the shared formula will be applied (A1:A10)
                int startRow = 0;   // Row 1 (zero‑based)
                int endRow   = 9;   // Row 10
                int column   = 0;   // Column A

                // Set the original formula in the first cell of the range
                Cell firstCell = cells[startRow, column];
                firstCell.Formula = "SUM(B1:B10)";

                // Apply the shared formula to the rest of the cells in the range
                for (int row = startRow + 1; row <= endRow; row++)
                {
                    // Share the formula from the first cell (A1)
                    cells[row, column].SetSharedFormula(firstCell.Formula, startRow, column);
                }

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }

            Console.WriteLine("Shared formulas applied to all workbooks.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
