// Title: Remove duplicate rows from an Excel worksheet based on columns A and B using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, checks for duplicate values in columns A and B, removes those rows, and saves the result. | Provide a C# snippet demonstrating how to invoke Worksheet.Cells.RemoveDuplicates for the first two columns with proper error handling. | Create a reusable C# method that accepts input and output file paths and column indexes, then eliminates duplicate rows using Aspose.Cells.
// Common Searches: asp.net aspose.cells remove duplicate rows based on two columns | c# example using Worksheet.Cells.RemoveDuplicates to clean Excel data | how to delete rows with identical values in columns A and B in an .xlsx file using Aspose.Cells | remove repeated rows from Excel workbook with Aspose.Cells .NET library
// Tags: Aspose.Cells RemoveDuplicates columns A B | C# deduplicate Excel worksheet Aspose.Cells | Worksheet.Cells.RemoveDuplicates .NET usage | clean Excel data duplicate rows Aspose.Cells | duplicate row elimination .xlsx C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // This C# program verifies the input.xlsx file, loads it with Aspose.Cells, removes rows that have identical values in columns A and B using Worksheet.Cells.RemoveDuplicates, and saves the cleaned workbook as output.xlsx while handling exceptions.
class RemoveDuplicateRows
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of rows (MaxDataRow is zero‑based)
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = 2; // Columns A and B (0‑based index)

            // Remove duplicate rows based on values in columns A and B
            // The last parameter (columnOffsets) is set to null because no offsets are needed
            sheet.Cells.RemoveDuplicates(0, 0, totalRows, totalColumns, false, null);

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Duplicates removed. Output saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
