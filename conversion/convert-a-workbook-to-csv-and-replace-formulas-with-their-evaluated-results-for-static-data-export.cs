// Title: Convert an Excel workbook to CSV with calculated values only using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, forces a full formula recalculation, removes all formulas, and writes the first sheet to a CSV file. | Show the steps to use Aspose.Cells in .NET to evaluate workbook formulas and export the resulting values as a CSV document.
// Common Searches: Aspose.Cells export first worksheet to CSV after evaluating formulas in C# | How to strip formulas from an Excel workbook before saving as CSV using .NET | C# convert .xlsx to .csv with only calculated values using Aspose.Cells | Save Excel workbook as CSV with values only, not formulas, Aspose.Cells library
// Tags: Aspose.Cells calculate formulas before CSV export | strip formulas Aspose.Cells C# | primary worksheet CSV export Aspose.Cells | Excel to CSV values-only .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // // Loads an Excel workbook, recalculates all formulas, removes the formulas, and saves the first worksheet as a CSV file using Aspose.Cells for .NET.
    public class WorkbookToCsvExport
    {
        public static void Run()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Path where the CSV file will be saved
            string csvPath = "output.csv";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook from the source file
                Workbook workbook = new Workbook(sourcePath);

                // Calculate all formulas in the workbook so that their results are up‑to‑date
                workbook.CalculateFormula();

                // Remove formulas from each worksheet, leaving only the calculated values
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Cells.RemoveFormulas();
                }

                // Save the workbook as CSV (the first worksheet is exported by default)
                workbook.Save(csvPath, SaveFormat.Csv);

                Console.WriteLine($"Workbook has been exported to CSV at: {csvPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                WorkbookToCsvExport.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
