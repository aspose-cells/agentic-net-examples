// Title: C# – Compute Running Total from a CSV and Export to XLSX with Aspose.Cells
// Description: Loads a CSV file into an Aspose.Cells workbook, calculates a cumulative sum for a specified numeric column, inserts the running total into a new column with a header, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | Import CSV | running total | cumulative sum | Excel export | SaveFormat.Xlsx | worksheet cells | data transformation | CSV to Excel conversion
// Common Searches: Aspose.Cells calculate running total from CSV | C# import CSV and add cumulative column using Aspose.Cells | How to create a running total column in Excel with Aspose.Cells .NET | Convert CSV to XLSX and compute cumulative sum Aspose.Cells | Aspose.Cells example for cumulative totals after CSV import
// Developer Intent: Add a cumulative‑total column to data imported from a CSV file and save the enhanced worksheet as an Excel workbook.
// Use Cases: Financial statements that show a running balance of revenue or expenses. | Sales dashboards displaying cumulative units sold or revenue over time. | Inventory logs that track a running stock balance after each transaction.
// AI Prompts: Generate C# code using Aspose.Cells to import a CSV, compute a running total for a given column, and save the workbook as XLSX with the totals in a new column. | Show an Aspose.Cells .NET example that adds a cumulative‑sum column after importing CSV data, handling non‑numeric rows gracefully. | Explain how to find the last used column in an Aspose.Cells worksheet and insert a new column for running totals.

using System;
using Aspose.Cells;

namespace AsposeCellsRunningTotalExample
{
    // Loads a CSV file into an Aspose.Cells workbook, calculates a cumulative sum for a specified numeric column, inserts the running total into a new column with a header, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input CSV file path and output Excel file path
            string csvPath = "input.csv";          // replace with actual CSV file location
            string outputPath = "output.xlsx";     // desired output file

            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import CSV data starting at cell A1 (row 0, column 0)
            // Using comma as delimiter, converting numeric data automatically
            cells.ImportCSV(csvPath, ",", true, 0, 0);   // lifecycle load

            // Determine the column that contains the numeric values for which we want a running total.
            // For this example we assume the numeric data is in column B (zero‑based index 1).
            int numericColumnIndex = 1;

            // Determine where to place the running total column: one column after the last used column.
            int totalColumnIndex = cells.MaxDataColumn + 1;

            // Optional: write a header for the running total column
            cells[0, totalColumnIndex].PutValue("Running Total");

            double runningSum = 0.0;

            // Iterate over all data rows (starting from row 1 to skip header)
            for (int row = 1; row <= cells.MaxDataRow; row++)
            {
                // Retrieve the value from the numeric column
                object val = cells[row, numericColumnIndex].Value;

                // Try to convert to double; if conversion fails, treat as zero.
                double number;
                if (val != null && double.TryParse(val.ToString(), out number))
                {
                    runningSum += number;
                }

                // Write the running total into the new column
                cells[row, totalColumnIndex].PutValue(runningSum);
            }

            // Save the workbook with the new column (lifecycle save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
