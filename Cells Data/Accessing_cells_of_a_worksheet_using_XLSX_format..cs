using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format is the default)
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain the Cells collection of the worksheet
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // Access cells using zero‑based row and column indexes
            // -------------------------------------------------
            // Cell A1 (row 0, column 0)
            cells[0, 0].PutValue("Hello World");

            // Cell C2 (row 1, column 2)
            cells[1, 2].PutValue(12345);

            // -------------------------------------------------
            // Access cells using the A1 style address
            // -------------------------------------------------
            // Cell B1
            cells["B1"].PutValue(DateTime.Now);

            // Cell C3 – set a formula that sums a range
            cells["C3"].Formula = "=SUM(A1:C2)";

            // -------------------------------------------------
            // Calculate all formulas in the workbook
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Read and display values from the cells
            // -------------------------------------------------
            Console.WriteLine("A1 (String) = " + cells[0, 0].StringValue);
            Console.WriteLine("C2 (Int)    = " + cells[1, 2].IntValue);
            Console.WriteLine("B1 (Date)   = " + cells["B1"].DateTimeValue);
            Console.WriteLine("C3 (Result) = " + cells["C3"].Value);

            // -------------------------------------------------
            // Save the workbook as an XLSX file
            // -------------------------------------------------
            workbook.Save("AccessCellsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}