// Title: Recalculate All Formulas in an Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file into an Aspose.Cells Workbook, triggers Workbook.CalculateFormula to evaluate every formula, optionally saves the refreshed workbook, and prints a confirmation to the console.
// Keywords: Aspose.Cells C# calculate formulas | Workbook.CalculateFormula example | recalculate Excel formulas .NET | update workbook calculations | save workbook after formula evaluation | load Excel file Aspose.Cells
// Common Searches: how to recalculate formulas in an Excel file with Aspose.Cells C# | Aspose.Cells calculate all formulas and save workbook | Workbook.CalculateFormula usage example | refresh Excel calculations programmatically .NET | batch recalc of Excel workbooks Aspose
// Developer Intent: Execute a full workbook formula refresh after loading an Excel file and optionally persist the changes.
// Use Cases: Refresh totals after programmatically modifying cell values before exporting a report. | Ensure financial models are up‑to‑date when generating PDFs or CSV extracts. | Automate nightly processing of multiple spreadsheets to guarantee accurate calculations before archiving.
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, runs Workbook.CalculateFormula, and saves the result. | Show how to limit CalculateFormula to a single worksheet in Aspose.Cells. | Give tips for optimizing performance when recalculating large workbooks with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCalculation
{
    // Loads an .xlsx file into an Aspose.Cells Workbook, triggers Workbook.CalculateFormula to evaluate every formula, optionally saves the refreshed workbook, and prints a confirmation to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook file
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the loaded workbook
            workbook.CalculateFormula();

            // (Optional) Save the workbook after calculation
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Display a confirmation message
            Console.WriteLine($"Formulas calculated and workbook saved to '{outputPath}'.");
        }
    }
}
