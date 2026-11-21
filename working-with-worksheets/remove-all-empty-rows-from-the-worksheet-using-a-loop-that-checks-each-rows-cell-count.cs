// Title: Remove empty rows from an Excel worksheet with Aspose.Cells for .NET by iterating rows and deleting blanks
// Common Searches: C# code to delete rows that have no data in an Excel file using Aspose.Cells | Aspose.Cells loop through worksheet rows and remove blank rows | How to programmatically clean empty rows from an .xlsx with Aspose.Cells .NET | Delete empty rows in Excel using Aspose.Cells Workbook API C# | Iterate backwards through rows to avoid index shift when deleting rows Aspose.Cells
// Tags: Aspose.Cells delete empty rows | Aspose.Cells iterate worksheet rows C# | Aspose.Cells backward row deletion .NET | Aspose.Cells check cell values before row removal | Aspose.Cells MaxDataRow usage

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, determines the last used row, iterates backward through each row, checks every cell for data, deletes rows that contain only empty cells, and saves the cleaned workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the index of the last used row
            int lastRowIndex = worksheet.Cells.MaxDataRow;

            // Iterate backwards to avoid index shift when deleting rows
            for (int rowIndex = lastRowIndex; rowIndex >= 0; rowIndex--)
            {
                bool isEmpty = true;

                // Check each cell in the current row for data
                for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
                {
                    if (!string.IsNullOrEmpty(worksheet.Cells[rowIndex, col].StringValue))
                    {
                        isEmpty = false;
                        break;
                    }
                }

                // Delete the row if it contains no data
                if (isEmpty)
                {
                    worksheet.Cells.DeleteRow(rowIndex);
                }
            }

            // Save the workbook after removing empty rows
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
