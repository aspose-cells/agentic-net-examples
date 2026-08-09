// Title: C# – Import CSV, sum a column with Aspose.Cells, and save as XLSX
// Description: A C# example that creates a workbook, imports data from a CSV file, applies a SUM formula to a numeric column, forces calculation, and saves the result in a new cell of an XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells CSV import C# | sum column Aspose.Cells | write formula C# | calculate workbook formula | save workbook as XLSX | C# read CSV and aggregate
// Common Searches: Aspose.Cells import CSV and sum column C# | C# calculate total of a CSV column with Aspose.Cells | write SUM formula to cell using Aspose.Cells | how to save calculated workbook as XLSX in C# | Aspose.Cells formula evaluation example
// Developer Intent: Load a CSV file, compute the total of a numeric column, place the result in a designated cell, and export the workbook as an Excel file.
// Use Cases: Financial reporting: import transaction CSVs and display the grand total in a summary sheet. | Sales aggregation: combine daily sales files and automatically calculate total sales per product column. | Data consolidation: import raw data sets, compute column totals, and generate ready‑to‑share XLSX reports.
// AI Prompts: Show how to limit the SUM to a specific range (e.g., B2:B100) instead of the whole column. | Demonstrate writing the sum result to a different worksheet within the same workbook. | Explain how to skip non‑numeric rows during CSV import with Aspose.Cells.

using System;
using Aspose.Cells;

// A C# example that creates a workbook, imports data from a CSV file, applies a SUM formula to a numeric column, forces calculation, and saves the result in a new cell of an XLSX file using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "data.csv";

        // Path where the resulting workbook will be saved
        string outputPath = "result.xlsx";

        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV data into the worksheet starting at cell A1
        // Parameters: file name, delimiter, convert numeric data, first row, first column
        cells.ImportCSV(csvPath, ",", true, 0, 0); // lifecycle rule: load via ImportCSV

        // Write a formula that sums the numeric column (e.g., column B)
        // The result will be placed in cell C1
        cells["C1"].Formula = "=SUM(B:B)"; // using Cell.Formula property

        // Calculate the formula so the sum value is materialized
        workbook.CalculateFormula(); // ensures the formula result is computed

        // Save the workbook with the calculated sum (lifecycle rule: save)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
