// Title: Convert Excel Formulas to Values and Export as CSV with Aspose.Cells for .NET
// Description: Loads an .xlsx workbook, calculates all formulas, removes the formula expressions, and saves the first worksheet as a CSV file that contains only the evaluated results.
// Keywords: Aspose.Cells | .NET | C# | calculate formulas | remove formulas | export to CSV | value‑only CSV | Workbook.Save CSV | RemoveFormulas method | CalculateFormula method | Excel to CSV conversion
// Common Searches: Aspose.Cells export CSV without formulas | C# convert Excel formulas to values before CSV | How to remove formulas in Aspose.Cells | Calculate all formulas Aspose.Cells .NET | Save workbook as CSV with values only | RemoveFormulas example C# | Export Excel to CSV using Aspose.Cells
// Developer Intent: Evaluate every formula, replace it with its result, then generate a CSV file.
// Use Cases: Create CSV reports from Excel templates where formulas must be resolved first. | Provide data extracts to systems that cannot process Excel formulas. | Automate batch conversion of multiple .xlsx files to value‑only CSV files.
// AI Prompts: Generate C# code that uses Aspose.Cells to calculate all formulas, remove them, and save the worksheet as CSV. | Explain the RemoveFormulas method and discuss alternatives for preserving formatting when exporting to CSV. | Show how to iterate through all worksheets, evaluate formulas, and combine their values into a single CSV output.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaToCsv
{
    // Loads an .xlsx workbook, calculates all formulas, removes the formula expressions, and saves the first worksheet as a CSV file that contains only the evaluated results.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path where the CSV file will be saved
            string csvFile = "output.csv";

            // Load the workbook from the source file
            Workbook workbook = new Workbook(sourceFile);

            // Calculate all formulas in the workbook so that each formula cell has a result
            workbook.CalculateFormula();

            // Replace every formula with its calculated value
            // This removes the formula text and keeps only the computed result
            workbook.Worksheets[0].Cells.RemoveFormulas();

            // Save the workbook as CSV; since formulas are already removed,
            // the CSV will contain only the calculated values
            workbook.Save(csvFile, SaveFormat.Csv);

            Console.WriteLine("Workbook has been converted to CSV with formulas replaced by values.");
        }
    }
}
