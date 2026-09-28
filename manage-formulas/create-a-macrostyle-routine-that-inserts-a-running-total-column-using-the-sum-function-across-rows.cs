// Title: Add a Running Total Column with Cumulative SUM Formulas Using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a new column and fill each cell with a cumulative SUM formula referencing the previous rows. | Show how to programmatically create a cumulative sum column in an Excel worksheet by applying a dynamic SUM range with Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# insert column and apply cumulative SUM formula for running total | How to create a running total column in Excel using Aspose.Cells .NET | C# example for adding a running total column with SUM function in Aspose.Cells workbook | Programmatically calculate cumulative sum per row with Aspose.Cells in .NET
// Tags: insert column with formula Aspose.Cells .NET | cumulative sum calculation Aspose.Cells | cumulative total column generation C# | Aspose.Cells workbook calculation example | dynamic range SUM in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, fills column A with values 1‑10, inserts column B, adds a "Running Total" header, and assigns each cell in column B a cumulative SUM formula that totals values from column A up to the current row. Formulas are calculated immediately and the workbook is saved as RunningTotal.xlsx.
class RunningTotalExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                // Values: 1, 2, 3, ... 10
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // Insert a new column for the running total (column B)
            // The second parameter indicates whether to copy the style (false = no copy)
            sheet.Cells.InsertColumn(1, false);

            // Add header for the running total column
            sheet.Cells[0, 1].PutValue("Running Total");

            // Apply SUM formula to each cell in the running total column
            // Formula: =SUM($A$2:A2) for row 2, =SUM($A$2:A3) for row 3, etc.
            for (int row = 1; row < 10; row++)
            {
                string formula = $"=SUM($A$2:A{row + 1})";
                sheet.Cells[row, 1].Formula = formula;
            }

            // Calculate the formulas immediately
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "RunningTotal.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
