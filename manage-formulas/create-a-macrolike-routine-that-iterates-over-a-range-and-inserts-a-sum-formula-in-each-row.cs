// Title: Macro‑style row‑wise SUM using Aspose.Cells SetSharedFormula in C#
// Description: Creates a new workbook, fills columns A and B with sample data, applies a shared row‑relative SUM formula ("=SUM(A1:B1)") to column C for a defined number of rows, recalculates the sheet, prints each row's values, and saves the file as RowSumFormulas.xlsx.
// Keywords: Aspose.Cells | C# | .NET | SetSharedFormula | shared formula | row‑wise SUM | apply formula to range | programmatic Excel calculation | save workbook with formulas | Excel automation
// Common Searches: Aspose.Cells apply same SUM formula to multiple rows | C# SetSharedFormula example | how to insert row‑relative formulas with Aspose.Cells | programmatically calculate Excel formulas in .NET | bulk add SUM formulas to a column using Aspose
// Developer Intent: Implement a macro‑like routine that iterates over a range and inserts a row‑relative SUM formula into each row with a single shared‑formula call.
// Use Cases: Generate a financial ledger where each row automatically totals adjacent columns. | Create a data‑entry template that sums two input columns for every new record without manual entry. | Speed up large worksheet processing by applying one shared SUM formula instead of setting formulas cell‑by‑cell.
// AI Prompts: Modify the code to sum columns A‑C instead of A‑B using SetSharedFormula. | Explain why SetSharedFormula is faster than assigning a formula to each cell individually. | Show how to read the existing row count from a worksheet and apply the shared SUM formula dynamically.

using System;
using Aspose.Cells;

namespace AsposeCellsMacroLikeSum
{
    // Creates a new workbook, fills columns A and B with sample data, applies a shared row‑relative SUM formula ("=SUM(A1:B1)") to column C for a defined number of rows, recalculates the sheet, prints each row's values, and saves the file as RowSumFormulas.xlsx.
    public class InsertRowSumFormulas
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Number of rows to process
                int totalRows = 10;

                // Populate sample data in columns A and B (0‑based indexes 0 and 1)
                for (int row = 0; row < totalRows; row++)
                {
                    cells[row, 0].PutValue(row + 1);          // A column: 1,2,3,...
                    cells[row, 1].PutValue((row + 1) * 10);   // B column: 10,20,30,...
                }

                // Target cell where the shared SUM formula will start (column C, index 2)
                Cell targetCell = cells[0, 2]; // C1

                // Apply the same row‑relative SUM formula to the whole column C
                targetCell.SetSharedFormula("=SUM(A1:B1)", totalRows, 1);

                // Recalculate the workbook so that all formulas are evaluated
                workbook.CalculateFormula();

                // Output the results to the console for verification
                for (int row = 0; row < totalRows; row++)
                {
                    Console.WriteLine($"Row {row + 1}: A={cells[row, 0].Value}, B={cells[row, 1].Value}, C (SUM)={cells[row, 2].Value}");
                }

                // Save the workbook (lifecycle rule: create → save)
                workbook.Save("RowSumFormulas.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            InsertRowSumFormulas.Run();
        }
    }
}
