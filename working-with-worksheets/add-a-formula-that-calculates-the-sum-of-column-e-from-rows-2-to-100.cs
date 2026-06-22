using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate column E (index 4) with sample data for rows 2 to 100
        for (int row = 1; row <= 99; row++) // zero‑based rows 1‑99 correspond to Excel rows 2‑100
        {
            sheet.Cells[row, 4].PutValue(row * 10); // example value
        }

        // Set the formula that sums E2:E100 in cell E101 (row index 100)
        sheet.Cells[100, 4].Formula = "=SUM(E2:E100)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the calculated sum to the console
        Console.WriteLine("Sum of E2:E100 = " + sheet.Cells[100, 4].Value);

        // Save the workbook to a file
        workbook.Save("SumColumnE.xlsx");
    }
}