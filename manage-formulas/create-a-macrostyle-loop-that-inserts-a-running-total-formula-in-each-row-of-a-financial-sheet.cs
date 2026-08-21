// Title: Add a Running Total Column with a Macro‑Style Loop Using Aspose.Cells for .NET (C#)
// Description: This example shows how to create a new workbook, populate it with financial data, and use a for‑loop to insert a cumulative SUM formula in column D for each row (e.g., =SUM($C$2:C5)). After assigning the formulas, Workbook.CalculateFormula() evaluates them and the workbook is saved as FinancialRunningTotal.xlsx.
// Keywords: Aspose.Cells | C# | .NET | running total formula | cumulative sum column | Excel formula loop | financial worksheet | macro style loop | calculate formulas | save workbook
// Common Searches: Aspose.Cells add cumulative sum column C# | insert running total formula each row Aspose.Cells | loop to set Excel formulas with Aspose.Cells | calculate running totals after populating data .NET | save workbook with running total column
// Developer Intent: Programmatically add a cumulative SUM formula to every data row of a financial sheet using Aspose.Cells for .NET.
// Use Cases: Generate cash‑flow statements that display the balance after each transaction. | Create expense reports where each line shows total spending to date. | Build dynamic financial dashboards that auto‑update running totals when new rows are added.
// AI Prompts: Provide C# code with Aspose.Cells that adds a cumulative SUM formula in column D for a table starting at row 2. | Show how to refresh running‑total formulas after inserting additional rows into an existing worksheet using Aspose.Cells. | Explain the role of Workbook.CalculateFormula in evaluating running totals and ensuring the results are written to the output file.

using System;
using Aspose.Cells;

namespace AsposeCellsRunningTotalDemo
{
    // This example shows how to create a new workbook, populate it with financial data, and use a for‑loop to insert a cumulative SUM formula in column D for each row (e.g., =SUM($C$2:C5)). After assigning the formulas, Workbook.CalculateFormula() evaluates them and the workbook is saved as FinancialRunningTotal.xlsx.
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

                // Header row
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Description");
                cells["C1"].PutValue("Amount");
                cells["D1"].PutValue("Running Total");

                // Sample financial data (Date, Description, Amount)
                object[,] data = new object[,]
                {
                    { new DateTime(2023, 1, 1), "Opening Balance", 1000.0 },
                    { new DateTime(2023, 1, 5), "Sales", 250.0 },
                    { new DateTime(2023, 1, 10), "Purchase", -150.0 },
                    { new DateTime(2023, 1, 15), "Service Income", 300.0 },
                    { new DateTime(2023, 1, 20), "Expense", -200.0 }
                };

                // Populate the sheet with the sample data starting from row 2 (index 1)
                int startRow = 1; // zero‑based index for the first data row
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    cells[startRow + i, 0].PutValue((DateTime)data[i, 0]); // Date
                    cells[startRow + i, 1].PutValue((string)data[i, 1]); // Description
                    cells[startRow + i, 2].PutValue((double)data[i, 2]); // Amount
                }

                // Insert running total formula in column D for each data row
                // Formula pattern: =SUM($C$2:C{currentRow})
                for (int row = startRow; row < startRow + data.GetLength(0); row++)
                {
                    // Build the formula string for the current row (Excel rows are 1‑based)
                    string formula = $"=SUM($C$2:C{row + 1})";
                    // Assign the formula to the cell
                    cells[row, 3].Formula = formula;
                }

                // Calculate all formulas so that the running totals are materialized
                workbook.CalculateFormula();

                // Save the workbook to a file
                workbook.Save("FinancialRunningTotal.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
