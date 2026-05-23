using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV file into the worksheet starting at cell A1 (row 0, column 0)
        // Using comma as delimiter and converting numeric strings to numbers
        string csvPath = "data.csv"; // path to your CSV file
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Assume the numeric column to sum is column B (zero‑based index 1)
        // Write a SUM formula into cell D1 (row 0, column 3)
        worksheet.Cells[0, 3].Formula = "=SUM(B:B)";

        // Calculate all formulas in the workbook so the sum value is evaluated
        workbook.CalculateFormula();

        // Save the workbook with the calculated sum
        workbook.Save("result.xlsx", SaveFormat.Xlsx);
    }
}