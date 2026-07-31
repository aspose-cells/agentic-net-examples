// Title: Aspose.Cells for .NET – Convert Excel formulas to static values and save as CSV
// Description: Load an Excel workbook, evaluate every formula, replace formulas with their results, and export the sheet(s) to a CSV file using Aspose.Cells.
// Keywords: Aspose.Cells CSV export | remove formulas Excel .NET | calculate workbook formulas | flatten formulas to values | C# export Excel to CSV | Aspose.Cells replace formulas | save workbook as CSV
// Common Searches: Aspose.Cells convert formulas to values before CSV | C# export Excel with calculated results to CSV | How to remove formulas when saving CSV with Aspose | Calculate all formulas then save as CSV using .NET | Flatten Excel formulas for CSV output
// Developer Intent: Replace every formula cell with its evaluated value and generate a CSV file.
// Use Cases: Produce CSV reports from template workbooks that contain formulas, ensuring only static data is delivered. | Prepare data extracts for systems that cannot interpret Excel formulas, requiring plain numbers in CSV. | Automate batch processing of multiple spreadsheets, flattening formulas before bulk CSV conversion.
// AI Prompts: Generate C# code with Aspose.Cells that calculates all formulas, removes them, and saves the workbook as CSV. | Explain step‑by‑step how to flatten formulas to values for CSV export using Aspose.Cells for .NET. | Show how to handle multiple worksheets when converting formula cells to static values before saving to CSV.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaToCsv
{
    // Load an Excel workbook, evaluate every formula, replace formulas with their results, and export the sheet(s) to a CSV file using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Replace formulas with their calculated values for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells.RemoveFormulas();
            }

            // Save the workbook as CSV (save rule)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook has been converted to CSV with formulas replaced by values: {csvPath}");
        }
    }
}
