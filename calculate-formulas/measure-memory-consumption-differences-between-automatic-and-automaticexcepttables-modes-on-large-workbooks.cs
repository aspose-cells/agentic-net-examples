// Title: C# benchmark: Memory usage of Automatic vs AutomaticExceptTable modes in Aspose.Cells
// Description: This example creates a 5,000‑row × 50‑column workbook, fills it with numeric data and a SUM formula per row, then measures the memory footprint of the Automatic and AutomaticExceptTable calculation modes. It forces garbage collection, clones the template, runs calculations, optionally saves the files, and reports the byte difference, helping developers decide which mode is more memory‑efficient for large workbooks.
// Keywords: Aspose.Cells memory benchmark | Automatic calculation mode | AutomaticExceptTable mode | C# workbook performance | large workbook memory usage | CalcModeType comparison | formula calculation memory impact | Aspose.Cells .NET profiling
// Common Searches: Aspose.Cells memory consumption Automatic vs AutomaticExceptTable | measure memory usage calculation mode Aspose.Cells C# | benchmark Aspose.Cells formula calculation memory | how to compare calculation modes memory Aspose.Cells | large workbook performance Aspose.Cells .NET
// Developer Intent: Find out how much memory each calculation mode (Automatic and AutomaticExceptTable) consumes when processing a large workbook with many formulas.
// Use Cases: Profile memory impact of different calculation modes to choose the most efficient setting for high‑load applications. | Validate that AutomaticExceptTable reduces RAM usage in workbooks with extensive formulas. | Create automated test suites that log memory consumption across CalcModeType values for various workbook sizes.
// AI Prompts: Generate a C# method that iterates over all CalcModeType values, measures memory usage for each, and writes the results to a CSV file. | Refactor the memory measurement code to avoid writing the workbook to disk while still forcing internal structures to materialize. | Explain how to interpret the memory difference output and recommend the optimal calculation mode for large, formula‑heavy workbooks.

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    // This example creates a 5,000‑row × 50‑column workbook, fills it with numeric data and a SUM formula per row, then measures the memory footprint of the Automatic and AutomaticExceptTable calculation modes. It forces garbage collection, clones the template, runs calculations, optionally saves the files, and reports the byte difference, helping developers decide which mode is more memory‑efficient for large workbooks.
    class Program
    {
        // Size of the test workbook
        const int RowCount = 5000;
        const int ColumnCount = 50;

        static void Main()
        {
            try
            {
                // Ensure Aspose.Cells license is set if needed
                // License license = new License();
                // license.SetLicense("Aspose.Cells.lic");

                Console.WriteLine("Generating large workbook data...");
                // Create a workbook and fill it with data and formulas
                Workbook wbTemplate = CreateLargeWorkbook();

                // Measure memory for Automatic calculation mode
                long memoryAutomatic = MeasureMemoryUsage(CalcModeType.Automatic, wbTemplate);
                // Measure memory for AutomaticExceptTable calculation mode
                long memoryAutomaticExceptTable = MeasureMemoryUsage(CalcModeType.AutomaticExceptTable, wbTemplate);

                Console.WriteLine();
                Console.WriteLine("Memory consumption (bytes):");
                Console.WriteLine($"Automatic               : {memoryAutomatic:N0}");
                Console.WriteLine($"AutomaticExceptTable    : {memoryAutomaticExceptTable:N0}");
                Console.WriteLine($"Difference (Automatic - AutomaticExceptTable): {memoryAutomatic - memoryAutomaticExceptTable:N0}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates a workbook with a large amount of data and simple formulas
        static Workbook CreateLargeWorkbook()
        {
            Workbook wb = new Workbook(); // using the provided constructor rule
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate numeric data
            for (int row = 0; row < RowCount; row++)
            {
                for (int col = 0; col < ColumnCount; col++)
                {
                    cells[row, col].PutValue(row + col);
                }
            }

            // Add a simple sum formula in the last column of each row
            int formulaCol = ColumnCount; // one column after the data
            for (int row = 0; row < RowCount; row++)
            {
                string range = $"A{row + 1}:{GetColumnName(ColumnCount - 1)}{row + 1}";
                cells[row, formulaCol].Formula = $"=SUM({range})";
            }

            // Save the template workbook to a temporary file (optional, demonstrates save rule)
            string templatePath = "LargeWorkbookTemplate.xlsx";
            try
            {
                wb.Save(templatePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save template workbook: {ex.Message}");
            }

            return wb;
        }

        // Measures memory usage for a given calculation mode
        static long MeasureMemoryUsage(CalcModeType mode, Workbook template)
        {
            // Force a full garbage collection before measurement
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long before = GC.GetTotalMemory(true);

            // Clone the template workbook
            Workbook wb = new Workbook();
            wb.Copy(template); // copy contents from the template

            // Set the desired calculation mode
            wb.Settings.FormulaSettings.CalculationMode = mode;

            // Trigger calculation (if any formulas exist)
            wb.CalculateFormula();

            // Optionally save to ensure all internal structures are materialized
            string fileName = $"Workbook_{mode}.xlsx";
            try
            {
                wb.Save(fileName, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook for mode {mode}: {ex.Message}");
            }

            // Force another garbage collection after operations
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            long after = GC.GetTotalMemory(true);
            return after - before;
        }

        // Helper to convert zero‑based column index to Excel column name (A, B, ..., AA, AB, ...)
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = string.Empty;
            do
            {
                name = letters[index % 26] + name;
                index = index / 26 - 1;
            } while (index >= 0);
            return name;
        }
    }
}
