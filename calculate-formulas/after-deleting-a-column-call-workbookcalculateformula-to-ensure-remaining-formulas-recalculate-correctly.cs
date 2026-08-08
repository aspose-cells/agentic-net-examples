// Title: Delete a Column and Recalculate Formulas with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to remove a worksheet column, update formula references, and force a full recalculation using Workbook.CalculateFormula in Aspose.Cells for .NET. The example prints the adjusted values and saves the workbook.
// Keywords: Aspose.Cells delete column C# | Workbook.CalculateFormula .NET | update formula references after column removal | recalculate Excel formulas programmatically | Aspose.Cells column deletion example
// Common Searches: how to delete a column and refresh formulas Aspose.Cells | Aspose.Cells recalculate after column removal | C# delete worksheet column and update formulas | Workbook.CalculateFormula usage after deleting column
// Developer Intent: Remove a specific column from a spreadsheet and ensure all dependent formulas are automatically updated and recalculated.
// Use Cases: Cleaning up generated reports by dropping unused columns while keeping totals accurate. | Adjusting financial models after removing a data series, with sums and averages updated instantly. | Automating spreadsheet maintenance tasks that require column deletions without breaking formula logic.
// AI Prompts: Provide C# code that deletes column B in an Aspose.Cells workbook and runs Workbook.CalculateFormula to refresh all formulas. | Explain why the updateReference flag is needed when calling Cells.DeleteColumn and how CalculateFormula completes the update. | Show an end‑to‑end example that deletes a column, recalculates dependent formulas, and saves the updated Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to remove a worksheet column, update formula references, and force a full recalculation using Workbook.CalculateFormula in Aspose.Cells for .NET. The example prints the adjusted values and saves the workbook.
    public class DeleteColumnAndRecalculateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Fill sample data in columns A, B, C
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["C1"].PutValue(30);

                // Add a formula that sums the three columns
                cells["D1"].Formula = "=SUM(A1:C1)";

                // Delete column B (index 1) and update references in formulas
                cells.DeleteColumn(1, true);

                // Recalculate all formulas after the column deletion
                workbook.CalculateFormula();

                // Display the updated values and formula
                Console.WriteLine("After deleting column B and recalculating:");
                Console.WriteLine($"A1 value: {cells["A1"].Value}");
                Console.WriteLine($"B1 value (original C1): {cells["B1"].Value}");
                Console.WriteLine($"D1 formula: {cells["D1"].Formula}");
                Console.WriteLine($"D1 value: {cells["D1"].Value}");

                // Save the modified workbook
                string outputPath = "DeleteColumnRecalcDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteColumnAndRecalculateDemo.Run();
        }
    }
}
