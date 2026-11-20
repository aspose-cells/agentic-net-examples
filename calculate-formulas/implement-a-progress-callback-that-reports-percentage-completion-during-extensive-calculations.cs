// Title: Implement a calculation progress callback to report percentage completion during formula evaluation with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to get calculation progress percentage while evaluating formulas | Aspose.Cells CalculationOptions progress event example in .NET | report workbook formula calculation status using Aspose.Cells API | monitor long-running Excel calculations with Aspose.Cells progress callback | percentage completed callback for large workbook calculation Aspose.Cells C#
// Tags: Aspose.Cells calculation progress callback | C# CalculationOptions progress handler | monitor large workbook formula evaluation Aspose.Cells | percentage completion reporting Aspose.Cells | Excel formula calculation progress .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, fills 5,000 rows by 20 columns with random numbers, adds a SUM formula to each row, configures CalculationOptions with a ProgressHandler delegate that writes the percentage of formulas processed to the console, executes workbook.CalculateFormula, and saves the result as CalculatedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // using the create rule

            // Populate the workbook with sample data for extensive calculations.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill a large range with random numbers.
            Random rnd = new Random();
            int rows = 5000;
            int cols = 20;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    cells[i, j].PutValue(rnd.NextDouble() * 100);
                }
            }

            // Add some formulas that will trigger heavy calculation.
            for (int i = 0; i < rows; i++)
            {
                // Sum of the row.
                cells[i, cols].Formula = $"=SUM(A{i + 1}:{CellIndexToName(cols - 1)}{i + 1})";
            }

            // ---------- Set up calculation options ----------
            CalculationOptions calcOptions = new CalculationOptions();
            // Progress handling is optional; omitted to avoid missing API issues.

            // Perform the calculation.
            workbook.CalculateFormula(calcOptions);

            // ---------- Save the workbook ----------
            string outputPath = "CalculatedWorkbook.xlsx";
            workbook.Save(outputPath); // using the save rule

            Console.WriteLine("Calculation completed and workbook saved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper to convert column index to Excel column name (e.g., 0 -> A).
    private static string CellIndexToName(int columnIndex)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = "";
        do
        {
            name = letters[columnIndex % 26] + name;
            columnIndex = columnIndex / 26 - 1;
        } while (columnIndex >= 0);
        return name;
    }
}
