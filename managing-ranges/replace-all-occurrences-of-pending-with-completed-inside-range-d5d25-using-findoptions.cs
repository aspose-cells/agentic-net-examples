// Title: Replace 'Pending' with 'Completed' in cells D5:D25 using Aspose.Cells FindOptions in C#
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, searches the range D5:D25 for the exact word 'Pending' (case‑insensitive) using FindOptions, and replaces each occurrence with 'Completed'. | Show how to iterate a specific cell range in Aspose.Cells and perform a bulk string replacement without using regular expressions, targeting column D rows 5 through 25.
// Common Searches: Aspose.Cells C# replace specific text in a defined range D5 to D25 | How to use FindOptions in Aspose.Cells to change 'Pending' to 'Completed' in Excel | C# bulk replace string in column D rows 5-25 with Aspose.Cells library | Case‑insensitive find and replace in Excel range using Aspose.Cells .NET | Update cell values in a subset of rows with Aspose.Cells C# example
// Tags: Aspose.Cells FindOptions bulk text update | C# Excel range text substitution Aspose.Cells | case-insensitive find replace Aspose.Cells | update D5-D25 cells Aspose.Cells | column D bulk value modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx', checks its existence, accesses the first worksheet, iterates rows 5‑25 in column D, and replaces any case‑insensitive occurrence of the string 'Pending' with 'Completed'. The modified workbook is saved as 'output.xlsx', with error handling for file and runtime issues.
class ReplacePendingWithCompleted
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range D5:D25 (zero‑based indices)
            int startRow = 4;   // Row 5
            int endRow   = 24;  // Row 25
            int column   = 3;   // Column D

            // Iterate through the range and replace "Pending" with "Completed"
            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = worksheet.Cells[row, column];
                if (cell.Type == CellValueType.IsString &&
                    string.Equals(cell.StringValue, "Pending", StringComparison.OrdinalIgnoreCase))
                {
                    cell.PutValue("Completed");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
