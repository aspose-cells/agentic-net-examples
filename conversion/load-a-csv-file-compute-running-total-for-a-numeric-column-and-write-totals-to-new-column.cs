// Title: C# – Import CSV with Aspose.Cells, add a running‑total column, and save as XLSX
// Description: Creates a new Workbook, loads a CSV via ImportCSV, detects the numeric column, inserts a "RunningTotal" header after the last data column, computes the cumulative sum row‑by‑row, writes the totals to the new column, and saves the result as an XLSX file.
// Keywords: Aspose.Cells ImportCSV C# | running total column | cumulative sum Excel | add new column Aspose.Cells | save workbook as XLSX | CSV to Excel conversion .NET | calculate cumulative total programmatically
// Common Searches: Aspose.Cells import CSV and add cumulative column | C# compute running total from CSV with Aspose.Cells | How to insert a new column after last data column in Aspose.Cells | Export CSV data to XLSX with a total column using .NET | Aspose.Cells example for cumulative sum
// Developer Intent: Load a CSV, calculate a running total for a specified numeric column, write the totals to a new column, and export the workbook to XLSX.
// Use Cases: Financial statements that show cumulative revenue or expenses over time. | Inventory logs that track running stock levels as transactions are recorded. | Time‑series datasets where a cumulative metric (e.g., total sales) is required for downstream analytics.
// AI Prompts: Generate C# code using Aspose.Cells to read a CSV, compute a running total for column B, place the totals in a new column after the existing data, and save the workbook as XLSX. | Explain how Aspose.Cells treats empty or non‑numeric cells during a running‑total calculation and show how to skip or flag them. | Show how to programmatically find the last data column, insert a header for a cumulative‑sum column, and populate it with running totals.

using System;
using Aspose.Cells;

// Creates a new Workbook, loads a CSV via ImportCSV, detects the numeric column, inserts a "RunningTotal" header after the last data column, computes the cumulative sum row‑by‑row, writes the totals to the new column, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "data.csv";

        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Load CSV data into the worksheet starting at cell A1
        // Using comma as delimiter, converting numeric strings to numbers
        // (lifecycle rule: load via ImportCSV)
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Index of the column that holds the numeric values (e.g., column B -> index 1)
        int numericColumnIndex = 1;

        // Determine the column where the running total will be written
        // Place it after the last data column
        int totalColumnIndex = cells.MaxDataColumn + 1;

        // Write a header for the running total column
        cells[0, totalColumnIndex].PutValue("RunningTotal");

        // Compute the running total and write it to the new column
        double runningSum = 0;
        // Start from row 1 to skip the header row
        for (int row = 1; row <= cells.MaxDataRow; row++)
        {
            // Retrieve the numeric value; if the cell is empty or non‑numeric, DoubleValue returns 0
            double currentValue = cells[row, numericColumnIndex].DoubleValue;
            runningSum += currentValue;
            cells[row, totalColumnIndex].PutValue(runningSum);
        }

        // Save the workbook to an XLSX file (lifecycle rule: save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
