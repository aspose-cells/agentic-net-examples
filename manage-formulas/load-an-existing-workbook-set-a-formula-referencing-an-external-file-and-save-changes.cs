// Title: Add and Calculate an External Workbook Formula with Aspose.Cells for .NET
// Description: Demonstrates how to load a primary workbook, create or load a source workbook, register an external link, set a formula that references a cell in the source file, recalculate using linked data sources, and save the updated workbook.
// Keywords: Aspose.Cells external link | C# Excel external reference | calculate linked formula Aspose | set formula referencing another workbook | save workbook after external formula | Aspose.Cells CalculationOptions | Excel external workbook example
// Common Searches: Aspose.Cells set formula to external Excel file C# | How to add external link in Aspose.Cells workbook | Calculate external workbook formulas with Aspose.Cells | Save workbook after linking to another workbook Aspose | C# example for external cell reference using Aspose.Cells
// Developer Intent: Create an external workbook reference, evaluate it, and persist the changes.
// Use Cases: Build a dashboard workbook that pulls live values from a data workbook. | Automate financial consolidation by linking multiple source files into a summary sheet. | Generate reports that require up‑to‑date figures from separate Excel files without manual copying.
// AI Prompts: Generate C# code with Aspose.Cells that adds an external link to "Data.xlsx", sets cell B5 in the main workbook to =[Data.xlsx]Sheet1!$B$5, recalculates, and saves the file. | Explain why CalculationOptions.LinkedDataSources must include the source workbook when evaluating external formulas in Aspose.Cells. | Provide error‑handling patterns for missing source workbooks when adding external links with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    // Demonstrates how to load a primary workbook, create or load a source workbook, register an external link, set a formula that references a cell in the source file, recalculate using linked data sources, and save the updated workbook.
    class Program
    {
        static void Main()
        {
            // Paths to the main workbook and the external workbook
            string mainWorkbookPath = "Main.xlsx";
            string externalWorkbookPath = "External.xlsx";

            try
            {
                // Ensure the main workbook exists; create a simple one if missing
                if (!File.Exists(mainWorkbookPath))
                {
                    var wb = new Workbook();
                    wb.Worksheets[0].Name = "Sheet1";
                    wb.Save(mainWorkbookPath);
                }

                // Ensure the external workbook exists; create a simple one with a value in A2 if missing
                if (!File.Exists(externalWorkbookPath))
                {
                    var extWb = new Workbook();
                    var extSheet = extWb.Worksheets[0];
                    extSheet.Name = "Sheet1";
                    extSheet.Cells["A2"].PutValue(123); // sample data
                    extWb.Save(externalWorkbookPath);
                }

                // Load the existing main workbook
                Workbook mainWorkbook = new Workbook(mainWorkbookPath);

                // Load the external workbook (used as a data source for the formula)
                Workbook externalWorkbook = new Workbook(externalWorkbookPath);

                // Ensure the main workbook knows its file name (important for external links)
                mainWorkbook.FileName = Path.GetFileName(mainWorkbookPath);

                // Add an external link entry for the external workbook (required for proper link handling)
                // Here we reference only "Sheet1" of the external file
                int linkIndex = mainWorkbook.Worksheets.ExternalLinks.Add(
                    externalWorkbookPath,
                    new string[] { "Sheet1" });

                // Set a formula in cell A1 that references a cell in the external workbook
                // Formula format: =[External.xlsx]Sheet1!$A$2
                Worksheet sheet = mainWorkbook.Worksheets[0];
                sheet.Cells["A1"].Formula = $"=[{Path.GetFileName(externalWorkbookPath)}]Sheet1!$A$2";

                // Calculate the formula using the external workbook as a linked data source
                CalculationOptions calcOptions = new CalculationOptions
                {
                    LinkedDataSources = new Workbook[] { externalWorkbook }
                };
                mainWorkbook.CalculateFormula(calcOptions);

                // Save the updated main workbook
                string outputPath = "Main_Updated.xlsx";
                mainWorkbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with external formula at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
