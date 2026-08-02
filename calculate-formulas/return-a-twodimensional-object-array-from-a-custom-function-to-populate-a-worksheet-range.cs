using System;
using Aspose.Cells;

namespace AsposeCellsCustomFunctionDemo
{
    class Program
    {
        // Custom function that returns a two‑dimensional object array.
        // In a real scenario this could be any logic that builds the data.
        static object[][] GetSampleData()
        {
            // Create a 3 × 2 array.
            return new object[][]
            {
                new object[] { 10, 20 },
                new object[] { 30, 40 },
                new object[] { 50, 60 }
            };
        }

        static void Main()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare input cells that the table formula will reference.
            // These cells are not used by the custom function itself,
            // but SetTableFormula requires them.
            cells[0, 0].Formula = "=A1+A2"; // dummy formula
            cells[0, 1].PutValue(1);        // row input values
            cells[1, 0].PutValue(2);        // column input values

            // Target cell where the data table will start.
            Cell targetCell = worksheet.Cells["C3"];

            // Get the two‑dimensional array from the custom function.
            object[][] values = GetSampleData();

            // Set a two‑variable data table using the returned array.
            // rowNumber = number of rows in the table (values.Length)
            // columnNumber = number of columns in the table (values[0].Length)
            // rowInputCell and columnInputCell are dummy references required by the API.
            targetCell.SetTableFormula(
                rowNumber: values.Length,
                columnNumber: values[0].Length,
                rowInputCell: "A1",
                columnInputCell: "A2",
                values: values);

            // Calculate formulas so that the table results are refreshed.
            workbook.CalculateFormula();

            // Optional: display the populated range in the console.
            Console.WriteLine("Populated data table (C3:D5):");
            for (int r = 2; r < 2 + values.Length; r++)          // rows 2‑based (C3 is row index 2)
            {
                for (int c = 2; c < 2 + values[0].Length; c++) // columns 2‑based (C column index 2)
                {
                    Console.Write($"{cells[r, c].Value}\t");
                }
                Console.WriteLine();
            }

            // Save the workbook.
            workbook.Save("CustomFunctionDataTable.xlsx");
        }
    }
}