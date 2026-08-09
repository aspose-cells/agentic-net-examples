// Title: C# – Export Excel to CSV with Formulas Evaluated Using Aspose.Cells
// Description: Load an .xlsx file with Aspose.Cells, calculate all formulas, replace each formula with its result via RemoveFormulas, and save the workbook as a static CSV file.
// Keywords: Aspose.Cells CSV export | C# Excel to CSV conversion | evaluate formulas Aspose.Cells | RemoveFormulas method | static CSV export | Workbook.Save CSV | calculate all formulas .NET
// Common Searches: export Excel to CSV with evaluated formulas Aspose.Cells | remove formulas before saving CSV C# | calculate workbook formulas Aspose.Cells then export | static CSV from Excel using Aspose.Cells .NET
// Developer Intent: Generate a CSV file from an Excel workbook where every formula is replaced by its calculated value.
// Use Cases: Create a data‑only CSV report from a workbook that contains complex calculations. | Automate batch conversion of multiple .xlsx files to CSV for downstream systems that require static values. | Prepare CSV files for import into databases or analytics tools without carrying over Excel formulas.
// AI Prompts: Write C# code with Aspose.Cells to open an .xlsx, evaluate all formulas, replace them with values, and save as CSV. | Explain the effect of Worksheet.Cells.RemoveFormulas on the CSV output produced by Aspose.Cells. | Provide performance tips for converting large workbooks to CSV while ensuring formulas are fully calculated.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an .xlsx file with Aspose.Cells, calculate all formulas, replace each formula with its result via RemoveFormulas, and save the workbook as a static CSV file.
    public class WorkbookToCsvExport
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(sourcePath);

                // Calculate all formulas in the workbook so that their results are up‑to‑date
                workbook.CalculateFormula();

                // Replace formulas with their calculated values for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // RemoveFormulas replaces each formula with its evaluated result
                    sheet.Cells.RemoveFormulas();
                }

                // Save the workbook as CSV (static data export, no formulas)
                workbook.Save("output.csv", SaveFormat.Csv);

                Console.WriteLine("Workbook has been exported to CSV with formulas evaluated.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
