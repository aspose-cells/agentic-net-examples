using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a sample formula in cell C1
            Cell targetCell = worksheet.Cells["C1"];
            targetCell.Formula = "=SUM(A1:B1)";

            // Obtain the exact textual representation of the formula
            // GetFormula(false, false) returns the formula in A1 notation without locale conversion
            string formulaText = targetCell.GetFormula(false, false);

            // Display the formula text
            Console.WriteLine("Exact formula text: " + formulaText);

            // Optionally, save the workbook (not required for formula extraction)
            workbook.Save("FormulaTextDemo.xlsx");
        }
    }
}