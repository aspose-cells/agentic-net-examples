// Title: C# – Import CSV, sum a column, and save as XLSX with Aspose.Cells
// Description: Demonstrates how to create a Workbook, import a CSV file using a comma delimiter, add a SUM formula for a numeric column, force calculation, and export the result to an XLSX file.
// Keywords: Aspose.Cells import CSV C# | sum column Excel Aspose.Cells | calculate formula workbook C# | CSV to XLSX conversion Aspose | ImportCSV method example | Excel formula calculation .NET
// Common Searches: Aspose.Cells import CSV and sum column | C# calculate total of CSV column with Aspose | How to add SUM formula after CSV import in .NET | Convert CSV to Excel and compute column total
// Developer Intent: Read a CSV file, compute the total of a numeric column, place the result in a new cell, and save the workbook as an XLSX file.
// Use Cases: Generate a sales summary by importing transaction CSV data and totaling the Amount column. | Create a financial dashboard that aggregates expense values from a CSV source into a single cell. | Automate conversion of raw data files to Excel while inserting a computed grand‑total for quick review.
// AI Prompts: Modify the example to sum only rows 2‑100 of column B. | Write the sum result to cell D5 and apply bold, background color, and number formatting. | Adapt the code for a semicolon‑delimited CSV and still calculate the column total.

using System;
using Aspose.Cells;

// Demonstrates how to create a Workbook, import a CSV file using a comma delimiter, add a SUM formula for a numeric column, force calculation, and export the result to an XLSX file.
class CsvSumExample
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "data.csv";

        // Path for the resulting Excel file
        string outputPath = "result.xlsx";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Write a formula that sums the entire numeric column (e.g., column B)
        // The result will appear in cell C1
        cells["C1"].Formula = "=SUM(B:B)";

        // Calculate the formula so the value is stored in the cell
        workbook.CalculateFormula();

        // Save the workbook to the specified output file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
