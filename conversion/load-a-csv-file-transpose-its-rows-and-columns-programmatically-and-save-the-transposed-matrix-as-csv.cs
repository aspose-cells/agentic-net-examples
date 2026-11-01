// Title: Transpose rows and columns of a CSV file using Aspose.Cells for .NET (C#) and export to a new CSV
// Common Searches: C# Aspose.Cells example to transpose CSV data and write to another CSV file | How to swap rows and columns in a CSV using Aspose.Cells workbook in .NET | Programmatically transpose a CSV matrix with Aspose.Cells and save as CSV in C# | Aspose.Cells transpose worksheet data from CSV and export result as CSV | Remove original sheet after transposing CSV with Aspose.Cells in C#
// Tags: Aspose.Cells transpose CSV worksheet | C# transpose rows to columns using Aspose.Cells | save transposed data as CSV with Aspose.Cells | remove original worksheet after CSV transpose Aspose.Cells | load CSV into workbook Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample loads 'input.csv' into an Aspose.Cells Workbook, creates a new worksheet, copies each cell to its transposed position (row ↔ column), optionally removes the original sheet, and saves the transposed matrix as 'output.csv' using SaveFormat.Csv.
class CsvTranspose
{
    static void Main()
    {
        // Load the source CSV file into a workbook
        Workbook workbook = new Workbook("input.csv");

        // Get the first worksheet (the CSV data is loaded into the first sheet)
        Worksheet sourceSheet = workbook.Worksheets[0];

        // Determine the used range dimensions
        int rowCount = sourceSheet.Cells.MaxDataRow + 1;      // total rows
        int colCount = sourceSheet.Cells.MaxDataColumn + 1;  // total columns

        // Add a new worksheet to hold the transposed data
        int transSheetIndex = workbook.Worksheets.Add();
        Worksheet transSheet = workbook.Worksheets[transSheetIndex];

        // Transpose: swap rows with columns
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                // Copy the value from source (r, c) to transposed (c, r)
                transSheet.Cells[c, r].Value = sourceSheet.Cells[r, c].Value;
            }
        }

        // Optional: remove the original sheet if you only want the transposed result
        workbook.Worksheets.RemoveAt(0);

        // Save the transposed matrix as a new CSV file
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
