// Title: Aspose.Cells for .NET: Loop through column A and set row‑specific multiplication formulas in column B (C#)
// Description: Creates a workbook, fills A1‑A10 with 1‑10, then uses a C# loop to write a formula in column B that multiplies each A‑cell by its row number (e.g., =A1*1, =A2*2). The workbook calculates the formulas, prints the results, and saves the file as MultiplicationFormulaDemo.xlsx.
// Keywords: Aspose.Cells C# example | set Excel formula programmatically | loop through rows Aspose.Cells | multiply cell by row number | calculate formulas .NET | save workbook with formulas | Excel automation Aspose | dynamic formula assignment | C# Excel API
// Common Searches: How to assign a formula to each row with Aspose.Cells | C# loop to create multiplication formulas in Excel | Aspose.Cells calculate formulas after setting them | Programmatically write row‑based formulas in .NET | Save Excel workbook after formula evaluation Aspose
// Developer Intent: Programmatically generate a row‑specific multiplication formula in column B that references column A and the current row index, evaluate the formulas, and persist the results in an Excel file.
// Use Cases: Bulk compute products of a value column and its row number without manual entry. | Generate dynamic reports where each row requires a custom calculation based on its position. | Automate data‑driven spreadsheets that need to be saved with pre‑calculated results for downstream systems.
// AI Prompts: Generate C# code using Aspose.Cells to fill column A with numbers 1‑100 and place =A{row}*{row} formulas in column B, then calculate and save the workbook. | Show how to modify the loop so column B multiplies column A by a constant factor or by another column while still using Aspose.Cells. | Explain how to detect the last used row in column A and apply the same multiplication formula to the entire used range with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, fills A1‑A10 with 1‑10, then uses a C# loop to write a formula in column B that multiplies each A‑cell by its row number (e.g., =A1*1, =A2*2). The workbook calculates the formulas, prints the results, and saves the file as MultiplicationFormulaDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample values (1 to 10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1..A10
        }

        // Loop through column A and set a multiplication formula in column B.
        // Formula: =A{row}*{row}
        for (int i = 0; i < 10; i++)
        {
            int rowNumber = i + 1; // Excel rows are 1‑based
            string formula = $"=A{rowNumber}*{rowNumber}";
            cells[i, 1].Formula = formula; // B column
        }

        // Calculate all formulas so that results are stored
        workbook.CalculateFormula();

        // Optional: display results in console
        for (int i = 0; i < 10; i++)
        {
            Console.WriteLine($"Row {i + 1}: A={cells[i, 0].Value}, B (A*row)={cells[i, 1].Value}");
        }

        // Save the workbook to a file
        workbook.Save("MultiplicationFormulaDemo.xlsx");
    }
}
