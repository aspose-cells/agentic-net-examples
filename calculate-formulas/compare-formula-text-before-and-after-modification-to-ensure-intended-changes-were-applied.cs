using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace AsposeCellsFormulaComparison
{
    public class FormulaComparisonDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // ---------- Create a workbook and set an initial formula ----------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Initial formula in cell A1
            worksheet.Cells["A1"].Formula = "=SUM(B1:B3)";

            // Populate referenced cells so the formula is valid
            worksheet.Cells["B1"].PutValue(10);
            worksheet.Cells["B2"].PutValue(20);
            worksheet.Cells["B3"].PutValue(30);

            // Change the formula to generate a revision entry
            worksheet.Cells["A1"].Formula = "=AVERAGE(B1:B3)";

            // Save the workbook (creates revision logs)
            string filePath = "FormulaRevisionDemo.xlsx";
            try
            {
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                return;
            }

            // ---------- Load the workbook and inspect revision logs ----------
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook revisionWorkbook;
            try
            {
                revisionWorkbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all revision logs
            foreach (RevisionLog log in revisionWorkbook.Worksheets.RevisionLogs)
            {
                foreach (Revision revision in log.Revisions)
                {
                    // Look for cell changes on A1
                    if (revision is RevisionCellChange cellChange && cellChange.CellName == "A1")
                    {
                        // Retrieve old and new formula texts
                        string oldFormula = cellChange.OldFormula;
                        string newFormula = cellChange.NewFormula;

                        // Output the formulas
                        Console.WriteLine($"Cell: {cellChange.CellName}");
                        Console.WriteLine($"Old Formula: {oldFormula}");
                        Console.WriteLine($"New Formula: {newFormula}");

                        // Compare the formulas
                        bool formulasAreEqual = string.Equals(oldFormula, newFormula, StringComparison.Ordinal);
                        Console.WriteLine($"Formulas are {(formulasAreEqual ? "identical" : "different")}.");

                        // Additional verification (optional)
                        if (!formulasAreEqual)
                        {
                            Console.WriteLine("Intended change was applied successfully.");
                        }
                    }
                }
            }
        }
    }
}