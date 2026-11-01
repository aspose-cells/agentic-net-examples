// Title: Scan an Excel workbook for INDIRECT formulas and assess their performance impact using Aspose.Cells for .NET
// Common Searches: how to find cells that use INDIRECT in an Excel file with Aspose.Cells C# | enumerate volatile formulas like INDIRECT in a .NET workbook using Aspose.Cells | detect and list INDIRECT function usage in Excel worksheets via C# Aspose.Cells API | load password protected Excel file and search for INDIRECT formulas with Aspose.Cells | measure performance impact of volatile functions when processing Excel with Aspose.Cells
// Tags: scan workbook for INDIRECT formulas Aspose.Cells | list volatile Excel functions C# | load password protected Excel file Aspose.Cells | calculate performance impact of INDIRECT .NET | iterate cells and read formulas Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIndirectFormulaScanner
{
    // The example loads an Excel workbook (prompting for a password when required), iterates through every worksheet and used cell, checks each formula for the INDIRECT function (case‑insensitive), records the sheet name, cell address, and formula, outputs the list, and notes that INDIRECT is a volatile function that can degrade calculation performance.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be analyzed.
            // Replace with the actual file path as needed.
            string workbookPath = "input.xlsx";

            // Verify that the file exists before attempting to load it.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found – \"{workbookPath}\"");
                return;
            }

            Workbook workbook = null;

            try
            {
                // Attempt to load the workbook without a password.
                workbook = new Workbook(workbookPath);
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                // Prompt for a password if the workbook is protected.
                Console.Write("The workbook is password‑protected. Enter password: ");
                string password = Console.ReadLine();

                try
                {
                    LoadOptions loadOptions = new LoadOptions { Password = password };
                    workbook = new Workbook(workbookPath, loadOptions);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to load workbook with provided password: {innerEx.Message}");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // List to hold information about cells that use the INDIRECT function.
            List<string> indirectFormulaCells = new List<string>();

            try
            {
                // Iterate through all worksheets in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all used cells in the worksheet.
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Check if the cell contains a formula.
                        if (cell.IsFormula)
                        {
                            // Determine if the formula uses the INDIRECT function (case‑insensitive).
                            if (cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the cell address and its formula.
                                indirectFormulaCells.Add($"{sheet.Name}!{cell.Name}: {cell.Formula}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while scanning formulas: {ex.Message}");
                return;
            }

            // Output the results.
            Console.WriteLine("Cells containing the INDIRECT function:");
            foreach (string entry in indirectFormulaCells)
            {
                Console.WriteLine(entry);
            }

            // Provide a brief note on the performance impact.
            Console.WriteLine();
            Console.WriteLine("Note: INDIRECT is a volatile function and may degrade calculation performance.");
        }
    }
}
