// Title: C# Loop to Insert a Running‑Total Formula in Every Row with Aspose.Cells
// Description: This example creates a new workbook, adds a simple financial table (Date, Description, Amount, Running Total), fills five sample rows, and uses a macro‑style loop to write a cumulative‑sum formula in column D for each data row (D2 = C2, Dn = D(n‑1)+Cn). After setting the formulas the workbook is recalculated and saved as FinancialRunningTotal.xlsx.
// Keywords: Aspose.Cells C# running total | cumulative sum formula Excel | programmatic Excel formulas loop | calculate running total Aspose | financial spreadsheet automation C#
// Common Searches: Aspose.Cells add running total column C# | loop to set Excel formulas with Aspose.Cells | cumulative sum in Excel using C# code | how to calculate running total programmatically | insert formula in each row Aspose.Cells
// Developer Intent: Programmatically place a running‑total formula in column D for every data row of a worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a sales ledger where each row shows the cumulative revenue. | Build a cash‑flow statement that updates the balance automatically as new entries are added. | Create budgeting worksheets that maintain progressive totals without manual formula entry.
// AI Prompts: Write C# code with Aspose.Cells that adds a running‑total column to an existing sheet, handling any number of rows. | Provide a macro‑style loop that assigns a cumulative‑sum formula to each row and then recalculates the workbook. | Explain how to recalculate all formulas after inserting running‑total formulas with Aspose.Cells and save the file.

using System;
using Aspose.Cells;

namespace AsposeCellsRunningTotalDemo
{
    // This example creates a new workbook, adds a simple financial table (Date, Description, Amount, Running Total), fills five sample rows, and uses a macro‑style loop to write a cumulative‑sum formula in column D for each data row (D2 = C2, Dn = D(n‑1)+Cn). After setting the formulas the workbook is recalculated and saved as FinancialRunningTotal.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ----- Set up a simple financial table -----
                // Header row
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Description");
                cells["C1"].PutValue("Amount");
                cells["D1"].PutValue("Running Total");

                // Sample data rows (rows 2‑6)
                string[] dates = { "2023-01-01", "2023-01-05", "2023-01-10", "2023-01-15", "2023-01-20" };
                string[] desc = { "Sale A", "Sale B", "Sale C", "Sale D", "Sale E" };
                double[] amounts = { 1200.50, 850.75, 430.00, 1025.30, 760.20 };

                for (int i = 0; i < dates.Length; i++)
                {
                    int row = i + 2; // Excel rows start at 1, our loop is zero‑based
                    cells[$"A{row}"].PutValue(dates[i]);
                    cells[$"B{row}"].PutValue(desc[i]);
                    cells[$"C{row}"].PutValue(amounts[i]);
                }

                // ----- Insert running‑total formula for each data row -----
                // Row 2 (first data row) – running total equals the first amount
                cells["D2"].Formula = "=C2";

                // Subsequent rows – running total = previous total + current amount
                for (int row = 3; row <= dates.Length + 1; row++)
                {
                    // Formula uses relative references to the previous total (D{row-1}) and current amount (C{row})
                    string formula = $"=D{row - 1}+C{row}";
                    cells[$"D{row}"].Formula = formula;
                }

                // Calculate all formulas so that the workbook contains the computed totals
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "FinancialRunningTotal.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
