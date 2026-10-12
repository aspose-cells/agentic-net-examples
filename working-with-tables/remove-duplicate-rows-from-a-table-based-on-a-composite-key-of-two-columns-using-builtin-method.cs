// Title: Remove duplicate rows from an Excel worksheet using Aspose.Cells RemoveDuplicates with a two‑column composite key (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells and removes rows that have duplicate values in both column A and column B. | Demonstrate how to invoke the duplicate‑removal routine to keep the first occurrence of each A‑B key while preserving the header row. | Show how to adjust the parameters for a dynamic range and save the cleaned workbook to a new file.
// Common Searches: Aspose.Cells C# remove duplicate rows based on columns A and B | How to use RemoveDuplicates for a composite key in an Excel file with Aspose.Cells | C# example of deleting duplicate Excel rows using Aspose.Cells Worksheet.Cells.RemoveDuplicates | Remove duplicate records from an Excel sheet using Aspose.Cells API with two key columns | Aspose.Cells RemoveDuplicates method parameters explanation for multi‑column keys
// Tags: Aspose.Cells duplicate row removal using composite key | Excel deduplication with two key columns C# | Aspose.Cells duplicate detection feature | C# Excel duplicate elimination Aspose.Cells | Composite key data cleaning Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an input.xlsx workbook, uses Aspose.Cells' RemoveDuplicates method to eliminate rows where the values in columns A and B (the composite key) are duplicated, and saves the cleaned data to output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the range that contains the data
                int firstRow = 0; // zero‑based index
                int totalRows = worksheet.Cells.MaxDataRow + 1; // total rows with data
                int firstColumn = 0; // column A
                int totalColumns = 2; // columns A and B form the composite key

                // Remove duplicate rows based on the composite key (columns A and B)
                // columnOffsets = null means use consecutive columns starting from firstColumn
                worksheet.Cells.RemoveDuplicates(firstRow, totalRows, firstColumn, totalColumns, false, null);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
