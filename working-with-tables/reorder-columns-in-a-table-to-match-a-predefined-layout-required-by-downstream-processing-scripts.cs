// Title: Reorder Excel worksheet columns by header names with Aspose.Cells in C#
// Common Searches: c# reorder excel columns based on header row using Aspose.Cells | aspocells copy specific columns to a new sheet by header name | move worksheet columns to match a predefined order in Excel with Aspose.Cells | load workbook and rearrange columns according to a list of headers in C# | example of using Aspose.Cells CopyColumn for column reordering
// Tags: aspocells copycolumn method | c# reorder excel columns aspocells | excel header based column ordering aspocells | aspocells create worksheet for reordered data | aspocells save workbook after column reordering

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample loads an existing Excel workbook, builds a dictionary that maps header names to source column indexes, adds a new worksheet, and uses Aspose.Cells' CopyColumn method to copy columns from the source sheet to the new sheet following a predefined header order, then saves the reordered workbook.
class Program
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the data to reorder is on the first worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Define the required column order by header names
            string[] requiredOrder = new string[] { "ID", "Name", "Date", "Amount" };

            // Create a new worksheet that will hold the reordered columns
            Worksheet targetSheet = workbook.Worksheets.Add("Reordered");

            // Build a map of header name -> source column index (first row is the header)
            Dictionary<string, int> headerToColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int maxColumn = sourceSheet.Cells.MaxColumn;
            for (int col = 0; col <= maxColumn; col++)
            {
                string header = sourceSheet.Cells[0, col].StringValue;
                if (!string.IsNullOrEmpty(header) && !headerToColumn.ContainsKey(header))
                {
                    headerToColumn[header] = col;
                }
            }

            // Copy columns to the target sheet following the required order
            for (int destCol = 0; destCol < requiredOrder.Length; destCol++)
            {
                string header = requiredOrder[destCol];
                if (headerToColumn.TryGetValue(header, out int srcCol))
                {
                    // Copy the whole column (including header) from source to target
                    // Correct overload: CopyColumn(sourceCells, sourceColumnIndex, destinationColumnIndex)
                    targetSheet.Cells.CopyColumn(sourceSheet.Cells, srcCol, destCol);
                }
                else
                {
                    // Header not found in source – column left blank
                }
            }

            // Save the workbook with the reordered layout
            workbook.Save(outputPath);
            Console.WriteLine($"Reordered workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
