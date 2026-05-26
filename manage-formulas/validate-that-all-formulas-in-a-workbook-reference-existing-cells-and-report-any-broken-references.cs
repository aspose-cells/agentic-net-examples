using System;
using Aspose.Cells;

namespace AsposeCellsFormulaValidator
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are parsed and calculated
            workbook.ParseFormulas(false);
            workbook.CalculateFormula();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only formula cells
                    if (!cell.IsFormula) continue;

                    // Get all precedents (references) appearing in the formula
                    ReferredAreaCollection precedents = cell.GetPrecedents();

                    // If there are no precedents, continue
                    if (precedents == null) continue;

                    // Examine each referenced area
                    foreach (ReferredArea area in precedents)
                    {
                        // Skip external links – they are handled differently
                        if (area.IsExternalLink) continue;

                        // Determine the worksheet that the reference points to
                        Worksheet refSheet = string.IsNullOrEmpty(area.SheetName)
                            ? sheet                                   // Same sheet if no sheet name specified
                            : workbook.Worksheets[area.SheetName];    // Look up by name

                        // If the referenced worksheet does not exist, report a broken reference
                        if (refSheet == null)
                        {
                            string refAddress = BuildReferenceAddress(area);
                            Console.WriteLine($"Broken reference in {sheet.Name}!{cell.Name}: {refAddress} (sheet not found)");
                            continue;
                        }

                        // Validate row/column indices (they must be non‑negative)
                        if (area.StartRow < 0 || area.StartColumn < 0 ||
                            (area.IsArea && (area.EndRow < area.StartRow || area.EndColumn < area.StartColumn)))
                        {
                            string refAddress = BuildReferenceAddress(area);
                            Console.WriteLine($"Broken reference in {sheet.Name}!{cell.Name}: {refAddress} (invalid range)");
                        }
                    }
                }
            }

            // Optionally save the workbook after validation (if any modifications were made)
            // workbook.Save("validated_output.xlsx");
        }

        // Helper method to build a human‑readable address string from a ReferredArea
        private static string BuildReferenceAddress(ReferredArea area)
        {
            string sheetPart = string.IsNullOrEmpty(area.SheetName) ? "" : $"{area.SheetName}!";
            string startCell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
            if (!area.IsArea) return $"{sheetPart}{startCell}";

            string endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
            return $"{sheetPart}{startCell}:{endCell}";
        }
    }
}