// Title: Import a CSV file, sum the second column with a SUM formula, and save the total to a new cell in an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to load data.csv, place a =SUM(B:B) formula in cell C1, calculate all formulas, and export the workbook as result.xlsx. | Write a C# program that calls Cells.ImportCSV to read a CSV, adds a SUM formula for column B, triggers workbook.CalculateFormula, and saves the file in XLSX format.
// Common Searches: Aspose.Cells C# import CSV and calculate column sum | How to add a SUM formula after importing CSV with Aspose.Cells .NET | Save CSV data as Excel with total of a column using Aspose.Cells | C# Aspose.Cells calculate formulas after ImportCSV | Sum values in second column of CSV and write result to Excel using Aspose.Cells
// Tags: ImportCSV method Aspose.Cells | SUM formula column B Aspose.Cells | Calculate workbook formulas Aspose.Cells | Save workbook as XLSX Aspose.Cells | aggregate numeric column CSV Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates how to import a CSV file into a workbook with Aspose.Cells for .NET, apply a SUM formula to column B, evaluate the formula, and save the result as an XLSX file.
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
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV data starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Place a formula in cell C1 that sums the entire second column (B)
        // The formula will automatically calculate the sum of all numeric values in column B
        cells["C1"].Formula = "=SUM(B:B)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to the specified output file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
