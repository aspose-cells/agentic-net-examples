// Title: Create a C# Aspose.Cells utility to list all INDIRECT formulas with their cell addresses and export to a new Excel report
// AI Prompts: Write a C# console program that loads a workbook with Aspose.Cells, iterates through every worksheet and cell, captures formulas containing the INDIRECT function, and writes each cell's full address and formula to a separate sheet in a new workbook. | Enhance the utility to also evaluate each INDIRECT formula and include its current value in the generated report alongside the address and formula text. | Add a command‑line flag that skips hidden worksheets during the scan, so only INDIRECT formulas from visible sheets are recorded in the output file.
// Common Searches: asp.net aspose.cells find cells that use INDIRECT function in an Excel file | c# generate report of formulas containing INDIRECT with Aspose.Cells | how to export addresses of indirect formulas to a new workbook using Aspose.Cells .NET | scan multiple worksheets for INDIRECT formulas and save results in Excel via C#
// Tags: Aspose.Cells scan workbook for specific formula | C# extract INDIRECT function usage | export formula addresses to Excel with Aspose | list cell references containing function in .NET | generate formula usage report Aspose.Cells

using System;
using Aspose.Cells;

namespace IndirectFormulaReport
{
    // The program loads "input.xlsx" with Aspose.Cells, traverses all worksheets and cells, identifies formulas that include the INDIRECT function, records each cell's full address and formula in a new worksheet, and saves the collection as "IndirectFormulasReport.xlsx".
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that will be scanned.
            string sourcePath = "input.xlsx";

            // Path where the generated report will be saved.
            string reportPath = "IndirectFormulasReport.xlsx";

            // Load the source workbook.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook that will hold the report.
            Workbook reportWorkbook = new Workbook();

            // Use the first worksheet for the report and give it a meaningful name.
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "Indirect Formulas";

            // Write header row.
            reportSheet.Cells[0, 0].PutValue("Cell Address");
            reportSheet.Cells[0, 1].PutValue("Formula");

            int reportRow = 1; // Start writing data from the second row.

            // Iterate through all worksheets in the source workbook.
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                Cells cells = ws.Cells;

                // Iterate through all cells that contain data/formulas.
                foreach (Cell cell in cells)
                {
                    // Check if the cell has a formula and if it contains the INDIRECT function.
                    if (cell.IsFormula && cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Build a full address that includes the sheet name.
                        string fullAddress = $"{ws.Name}!{cell.Name}";

                        // Write the address and the formula to the report sheet.
                        reportSheet.Cells[reportRow, 0].PutValue(fullAddress);
                        reportSheet.Cells[reportRow, 1].PutValue(cell.Formula);
                        reportRow++;
                    }
                }
            }

            // Save the report workbook.
            reportWorkbook.Save(reportPath);
        }
    }
}
