using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to the workbooks
                string mainWorkbookPath = "Main.xlsx";
                string externalWorkbookPath = "External.xlsx";

                // Verify that the files exist before loading
                if (!File.Exists(mainWorkbookPath))
                {
                    Console.WriteLine($"Main workbook not found: {mainWorkbookPath}");
                    return;
                }
                if (!File.Exists(externalWorkbookPath))
                {
                    Console.WriteLine($"External workbook not found: {externalWorkbookPath}");
                    return;
                }

                // Load the main workbook
                Workbook mainWorkbook = new Workbook(mainWorkbookPath);

                // Add an external link (full path is recommended)
                int externalLinkIndex = mainWorkbook.Worksheets.ExternalLinks.Add(
                    Path.GetFullPath(externalWorkbookPath),
                    new string[] { "Sheet1" });

                // Set a formula that references the external workbook
                Worksheet sheet = mainWorkbook.Worksheets[0];
                // Formula format: =[External.xlsx]Sheet1!$A$2
                sheet.Cells["A1"].Formula = $"=[{Path.GetFileName(externalWorkbookPath)}]Sheet1!$A$2";

                // Load the external workbook so Aspose.Cells can resolve the reference
                Workbook externalWorkbook = new Workbook(externalWorkbookPath);
                externalWorkbook.FileName = Path.GetFileName(externalWorkbookPath);

                // Update the main workbook with the external data source
                mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

                // Recalculate formulas (optional but ensures the result is up‑to‑date)
                mainWorkbook.CalculateFormula();

                // Save the updated main workbook
                string updatedPath = "Main_Updated.xlsx";
                mainWorkbook.Save(updatedPath);

                Console.WriteLine($"Workbook saved with external formula at: {updatedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}