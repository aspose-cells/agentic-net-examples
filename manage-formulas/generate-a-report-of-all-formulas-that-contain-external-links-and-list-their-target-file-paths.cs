// Title: Report Excel formulas with external links and their file paths using Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans all worksheets and cells, identifies formulas that contain external workbook references, extracts the precedents, and prints the worksheet name, cell address, and linked file path to the console.
// Keywords: Aspose.Cells external link detection | C# list Excel formulas referencing other workbooks | retrieve external file paths from formulas .NET | audit workbook for cross‑workbook references | enumerate formula precedents with external links | detect external references in Excel using Aspose | report external link cells Aspose.Cells
// Common Searches: How to find formulas that link to other workbooks with Aspose.Cells | C# code to list external file names used in Excel formulas | Aspose.Cells get precedents of a formula that are external links | Extract external workbook references from an Excel file .NET | Report cells containing external links in a spreadsheet
// Developer Intent: Identify every formula that references an external workbook and output its sheet, cell address, and the target file path.
// Use Cases: Perform a compliance audit of a workbook’s external data dependencies. | Generate documentation of cross‑workbook links for release management. | Create a validation tool that flags cells with external references for security review.
// AI Prompts: Write C# code with Aspose.Cells that exports all external‑link formulas to a CSV file. | Show how to modify the loop to also capture the exact range address of each external reference. | Provide a method that returns a collection of objects containing worksheet name, cell address, external file name, and referenced range.

using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkReport
{
    // Loads an Excel workbook, scans all worksheets and cells, identifies formulas that contain external workbook references, extracts the precedents, and prints the worksheet name, cell address, and linked file path to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that needs to be analyzed.
            // Replace with the actual file path as required.
            string workbookPath = "input.xlsx";

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Cells cells = worksheet.Cells;

                // Iterate through all cells that contain formulas.
                foreach (Cell cell in cells)
                {
                    // Check if the cell is a formula and contains an external link.
                    if (cell.IsFormula && cell.ContainsExternalLink)
                    {
                        // Get all references (precedents) used in the formula.
                        ReferredAreaCollection precedents = cell.GetPrecedents();

                        // Some formulas may not return precedents (null), guard against it.
                        if (precedents != null)
                        {
                            foreach (ReferredArea area in precedents)
                            {
                                // Identify external links among the precedents.
                                if (area.IsExternalLink)
                                {
                                    // Output the worksheet name, cell address, and external file path.
                                    Console.WriteLine($"Worksheet: {worksheet.Name}, Cell: {cell.Name}, External File: {area.ExternalFileName}");
                                }
                            }
                        }
                    }
                }
            }

            // Optionally, keep the workbook unchanged; no save operation is required for reporting.
        }
    }
}
