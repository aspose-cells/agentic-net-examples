using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VariableArrayMarkerOneDimensionalTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set initial input value and formula
                cells["A1"].PutValue(10);
                Cell formulaCell = cells["B1"];
                formulaCell.Formula = "=A1*2";

                // Input values for the one‑dimensional data table
                double[] inputValues = new double[] { 5, 10, 15, 20, 25 };

                // Convert to object[][] required by SetTableFormula
                object[][] tableValues = new object[inputValues.Length][];
                for (int i = 0; i < inputValues.Length; i++)
                {
                    tableValues[i] = new object[] { inputValues[i] };
                }

                // Create the data table (row input)
                formulaCell.SetTableFormula(
                    rowNumber: inputValues.Length,
                    columnNumber: 1,
                    inputCell: "A1",
                    isRowInput: true,
                    values: tableValues);

                // Recalculate formulas so results are materialized
                workbook.CalculateFormula();

                // Display generated table values
                Console.WriteLine("Generated one‑dimensional data table (B1:C5):");
                for (int r = 0; r < inputValues.Length; r++)
                {
                    Console.WriteLine($"Row {r + 1}: Result={cells[r, 1].Value}, Input={cells[r, 2].Value}");
                }

                // Save the workbook
                string outputPath = "VariableArrayMarkerOneDimensionalTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            VariableArrayMarkerOneDimensionalTable.Run();
        }
    }
}