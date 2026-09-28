// Title: Import a CSV file, compute a running total for a numeric column, and export the worksheet to XLSX using Aspose.Cells for .NET
// AI Prompts: Create C# code that uses Aspose.Cells to import a CSV file, calculate a cumulative sum for a specified numeric column, write the running totals into a new column, and save the workbook as an XLSX file. | Adapt the example to accept a custom delimiter and output the updated data back to a CSV file while still using Aspose.Cells for the calculations.
// Common Searches: Aspose.Cells C# how to add a cumulative total column to data imported from CSV | C# compute running total for column B after importing CSV with Aspose.Cells | Save modified worksheet as XLSX after processing CSV using Aspose.Cells .NET | Determine last data row in Aspose.Cells when working with imported CSV data
// Tags: import csv Aspose.Cells C# | calculate running total column Aspose.Cells | write cumulative values to new worksheet column | save workbook as xlsx Aspose.Cells | retrieve last data row Aspose.Cells

using System;
using Aspose.Cells;

// // Uses Aspose.Cells to import data.csv, computes a running total of the numeric values in column B, writes the totals to column C, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Path to the source CSV file
        string csvPath = "data.csv";

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Assume the numeric column to total is column B (zero‑based index 1)
        // The running total will be written to column C (zero‑based index 2)
        int headerRow = 0;               // header is in the first row
        int dataStartRow = headerRow + 1;
        double runningTotal = 0;

        // Determine the last row that contains data
        int lastDataRow = cells.MaxDataRow;

        // Write a header for the new total column
        cells[headerRow, 2].PutValue("RunningTotal");

        // Compute running total row by row
        for (int row = dataStartRow; row <= lastDataRow; row++)
        {
            // Retrieve the numeric value from column B; if the cell is empty, treat it as 0
            double currentValue = cells[row, 1].DoubleValue;
            runningTotal += currentValue;

            // Store the running total in column C of the same row
            cells[row, 2].PutValue(runningTotal);
        }

        // Save the workbook to an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
