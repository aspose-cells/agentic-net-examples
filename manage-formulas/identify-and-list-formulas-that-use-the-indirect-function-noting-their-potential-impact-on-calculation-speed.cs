// Title: Detect and list all INDIRECT formulas in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, scans every worksheet, and returns a collection of cell addresses and formulas where the formula contains the INDIRECT function. | Create a method in C# that groups found INDIRECT formulas by worksheet and prints the count of such formulas per sheet. | Generate a reusable utility that identifies any volatile functions (e.g., INDIRECT, OFFSET) in an Excel file and outputs their locations and formulas.
// Common Searches: aspocells c# find cells with INDIRECT function in excel | how to list volatile formulas like INDIRECT using Aspose.Cells .NET | detect performance‑heavy INDIRECT formulas in large Excel workbooks programmatically | C# scan workbook for indirect references and get cell addresses | Aspose.Cells example to enumerate formulas containing INDIRECT
// Tags: Aspose.Cells enumerate INDIRECT formulas | C# detect volatile Excel functions | identify indirect references in Excel via .NET | optimize workbook performance by reducing volatile formulas | list cells containing INDIRECT using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIndirectFinder
{
    // The example loads an Excel file with Aspose.Cells, iterates through each worksheet's used range, captures the sheet name, cell address, and formula for any cell whose formula includes the INDIRECT function, stores these details, and prints them while noting that INDIRECT is a volatile function that can degrade calculation speed in large workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be analyzed
            string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: File not found - {inputFile}");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(inputFile);

                // List to hold information about cells using INDIRECT
                List<string> indirectFormulas = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // Iterate through each cell in the used range
                    foreach (Cell cell in usedRange)
                    {
                        // Check if the cell contains a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Identify formulas that use the INDIRECT function (case‑insensitive)
                            if (cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the cell address, sheet name, and formula
                                string info = $"Sheet: {sheet.Name}, Cell: {cell.Name}, Formula: {cell.Formula}";
                                indirectFormulas.Add(info);
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Formulas using INDIRECT function:");
                foreach (string entry in indirectFormulas)
                {
                    Console.WriteLine(entry);
                }

                // Note on performance impact:
                // The INDIRECT function is volatile; it forces Excel to recalculate the formula
                // whenever any cell changes, which can significantly degrade calculation speed
                // in large workbooks. Identifying and minimizing its use can improve performance.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
