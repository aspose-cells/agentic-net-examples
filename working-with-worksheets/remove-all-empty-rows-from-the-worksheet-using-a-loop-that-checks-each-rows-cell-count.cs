// Title: Remove blank rows from an Excel worksheet with Aspose.Cells for .NET by iterating rows and checking cell values
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, scans each row from the bottom up, and deletes rows that contain only empty cells. | Demonstrate how to use Worksheet.Cells.DeleteRow in a loop to clean an Excel sheet by removing rows without any data with Aspose.Cells.
// Common Searches: C# Aspose.Cells delete rows that have no data in an Excel file | how to programmatically remove empty rows from a worksheet using Aspose.Cells .NET | iterate through rows bottom-up and delete blank rows with Aspose.Cells | Aspose.Cells example for cleaning up blank rows in a workbook
// Tags: Aspose.Cells delete empty rows | Worksheet.Cells.DeleteRow usage | scan rows for null values Aspose.Cells | clean Excel worksheet Aspose.Cells .NET | max data row loop Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads input.xlsx with Aspose.Cells, determines the last data row, iterates from that row upward, checks each cell in the row for a non‑null value, deletes rows that contain no data, and saves the cleaned workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int maxRow = sheet.Cells.MaxDataRow;

            // Iterate from the bottom up to avoid index shifting when deleting rows
            for (int row = maxRow; row >= 0; row--)
            {
                // Use the sheet's maximum data column as an upper bound for scanning
                int maxCol = sheet.Cells.MaxDataColumn;
                if (maxCol < 0)
                {
                    // No data in the sheet; delete the row
                    sheet.Cells.DeleteRow(row);
                    continue;
                }

                bool hasData = false;

                // Scan cells in the row to see if any cell has a non‑null value
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Value != null)
                    {
                        hasData = true;
                        break;
                    }
                }

                // If the row has no data, delete it
                if (!hasData)
                {
                    sheet.Cells.DeleteRow(row);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
