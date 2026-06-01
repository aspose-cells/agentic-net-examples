using System;
using Aspose.Cells;

namespace AsposeCellsSharedFormulaDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in columns A and B (rows 1 to 5)
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue(row + 1);          // Column A: 1,2,3,4,5
                cells[row, 1].PutValue((row + 1) * 10);   // Column B: 10,20,30,40,50
            }

            // Define the shared formula that adds values from column A and B
            string sharedFormula = "=A1+B1";

            // Set the shared formula starting at cell C1, spanning 5 rows and 3 columns (C1:E5)
            Cell startCell = cells[0, 2]; // C1
            FormulaParseOptions parseOptions = new FormulaParseOptions(); // default options
            startCell.SetSharedFormula(sharedFormula, 5, 3, parseOptions);

            // Display the propagated formulas for verification
            Console.WriteLine("C1 formula: " + cells[0, 2].Formula);
            Console.WriteLine("E5 formula: " + cells[4, 4].Formula);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated values for the range C1:E5
            Console.WriteLine("\nCalculated values (C1:E5):");
            for (int row = 0; row < 5; row++)
            {
                for (int col = 2; col < 5; col++) // columns C (2) to E (4)
                {
                    Console.WriteLine($"Cell {cells[row, col].Name} = {cells[row, col].Value}");
                }
            }

            // Verify that the cells are part of a shared formula
            Console.WriteLine("\nShared formula flags:");
            Console.WriteLine($"C1 IsSharedFormula: {cells[0, 2].IsSharedFormula}");
            Console.WriteLine($"E5 IsSharedFormula: {cells[4, 4].IsSharedFormula}");

            // Save the workbook to a file
            workbook.Save("SharedFormulaResult.xlsx");
            Console.WriteLine("\nWorkbook saved as 'SharedFormulaResult.xlsx'.");
        }
    }
}