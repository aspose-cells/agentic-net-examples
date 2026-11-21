// Title: How to export only visible rows from an Excel worksheet to a tab‑delimited text file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export visible rows to tab delimited text file | skip hidden rows when saving Excel worksheet as TXT with Aspose.Cells | write Excel data to .txt ignoring hidden rows in .NET | convert first worksheet of XLSX to tab‑separated values using Aspose.Cells | C# code to export Excel visible rows to text file with Aspose.Cells
// Tags: visible rows export to TSV using Aspose.Cells | hidden row filtering Aspose.Cells .NET | worksheet to text file conversion C# | Aspose.Cells tab‑delimited output example | Excel to plain text export ignoring hidden rows

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, determines the used range, and iterates through each row. It skips rows marked as hidden, gathers cell values, and writes them as tab‑separated lines to an output TXT file, handling errors and ensuring the source file exists.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string excelPath = "input.xlsx";
        string txtPath = "output.txt";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet (change index if another sheet is needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of rows and columns
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            // Open a StreamWriter to write the visible data to a TXT file
            using (StreamWriter writer = new StreamWriter(txtPath))
            {
                // Iterate through each row in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    // Skip rows that are hidden
                    // Use GetRow to obtain the Row object safely
                    Row currentRow = sheet.Cells.GetRow(row);
                    if (currentRow != null && currentRow.IsHidden)
                        continue;

                    // Collect cell values for the current row
                    string[] values = new string[maxColumn + 1];
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        values[col] = cell.Value?.ToString() ?? string.Empty;
                    }

                    // Write the row to the TXT file, using tab as a delimiter
                    writer.WriteLine(string.Join("\t", values));
                }
            }

            Console.WriteLine($"Data successfully exported to \"{txtPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
