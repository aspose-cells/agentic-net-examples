// Title: How to transpose rows and columns of a CSV file with Aspose.Cells for .NET and export the transposed matrix as CSV
// Common Searches: C# Aspose.Cells transpose CSV data matrix and save as new CSV file | programmatically swap rows and columns in a CSV using Aspose.Cells workbook | load CSV into Aspose.Cells, transpose worksheet, export to CSV in .NET | Aspose.Cells example for matrix transpose of CSV data in C# | convert CSV to transposed CSV using Aspose.Cells SaveFormat.Csv
// Tags: transpose CSV matrix using Aspose.Cells | Aspose.Cells read CSV workbook C# | swap worksheet rows columns Aspose.Cells | export transposed data to CSV with Aspose.Cells | Aspose.Cells SaveFormat.Csv example

using System;
using Aspose.Cells;

// The sample loads an input CSV into an Aspose.Cells Workbook, copies its cell values into a two‑dimensional array, writes the values to a new workbook with row and column indices swapped, and saves the transposed matrix as a CSV file using SaveFormat.Csv.
class CsvTranspose
{
    static void Main()
    {
        // Load the source CSV file into a workbook.
        // Aspose.Cells automatically detects CSV format.
        Workbook sourceWorkbook = new Workbook("input.csv");

        // Get the first worksheet (the CSV data is stored here).
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        Cells sourceCells = sourceSheet.Cells;

        // Determine the used range dimensions.
        int maxRow = sourceCells.MaxDataRow;   // zero‑based index of last row with data
        int maxCol = sourceCells.MaxDataColumn; // zero‑based index of last column with data

        int rowCount = maxRow + 1;
        int colCount = maxCol + 1;

        // Read the original data into a 2‑dimensional array.
        object[,] originalData = new object[rowCount, colCount];
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                originalData[r, c] = sourceCells[r, c].Value;
            }
        }

        // Create a new workbook for the transposed data.
        Workbook transposedWorkbook = new Workbook();
        Worksheet transposedSheet = transposedWorkbook.Worksheets[0];
        Cells transposedCells = transposedSheet.Cells;

        // Write transposed values (swap rows and columns).
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                // Destination row becomes original column, destination column becomes original row.
                transposedCells[c, r].PutValue(originalData[r, c]);
            }
        }

        // Save the transposed matrix as a CSV file.
        transposedWorkbook.Save("output_transposed.csv", SaveFormat.Csv);
    }
}
