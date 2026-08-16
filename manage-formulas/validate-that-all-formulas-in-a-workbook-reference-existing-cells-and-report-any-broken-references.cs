// Title: C# Aspose.Cells – Detect Missing Worksheet References in Excel Formulas
// Description: Loads an Excel file, parses all formulas, walks each worksheet and cell, extracts precedent areas, skips external links, and reports any formula that points to a non‑existent sheet.
// Keywords: Aspose.Cells formula validation | C# Excel broken sheet reference | detect missing worksheet in formula | Excel formula precedents check | programmatic formula integrity .NET | Excel data quality validation | global Excel automation
// Common Searches: Aspose.Cells find formulas referencing non‑existent sheets | C# code to validate Excel formula references | how to check broken worksheet links in Excel using .NET | detect missing sheet references in Excel formulas programmatically | validate Excel workbook formulas with Aspose.Cells
// Developer Intent: Identify every formula cell that references a worksheet that does not exist in the current workbook and surface those errors.
// Use Cases: Run the validator before distributing a workbook to guarantee all formulas resolve correctly. | Integrate the check into CI/CD pipelines for Excel‑based reporting solutions. | Generate a log of cells with invalid sheet references for audit trails and automated correction.
// AI Prompts: Create a method that returns a list of Cell objects with broken sheet references instead of writing to the console. | Enhance the validator to also flag references to rows or columns outside the used range of an existing sheet. | Write a unit test using Aspose.Cells that confirms the validator catches a formula pointing to a missing worksheet.

using System;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsFormulaValidator
{
    // Loads an Excel file, parses all formulas, walks each worksheet and cell, extracts precedent areas, skips external links, and reports any formula that points to a non‑existent sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure all formulas are parsed before analysis
            workbook.ParseFormulas(false);

            // Iterate through each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                Cells cells = worksheet.Cells;

                // Iterate through each cell in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only formula cells
                    if (cell.IsFormula)
                    {
                        // Get all references (precedents) used in the formula
                        ReferredAreaCollection precedents = cell.GetPrecedents();

                        if (precedents == null) continue;

                        foreach (ReferredArea area in precedents)
                        {
                            // Skip external links; they are not validated here
                            if (area.IsExternalLink) continue;

                            // Verify that the referenced sheet exists in the workbook
                            bool sheetExists = workbook.Worksheets.Any(ws => ws.Name.Equals(area.SheetName, StringComparison.OrdinalIgnoreCase));

                            if (!sheetExists)
                            {
                                Console.WriteLine($"Broken reference in cell {cell.Name} (Formula: {cell.Formula}) -> Sheet '{area.SheetName}' does not exist.");
                            }
                        }
                    }
                }
            }

            // Optionally, save the workbook after validation (if any modifications were made)
            // workbook.Save("validated_output.xlsx");
        }
    }
}
