using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (XLSX) with default load options.
            // ParsingFormulaOnOpen is true by default, ensuring formulas are parsed during load.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true;

            // Replace "input.xlsx" with the path to your source workbook.
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Access the first worksheet (you can change the index as needed).
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Duplicate an existing named range.
            // ------------------------------------------------------------
            // Assume there is an existing named range called "MyRange".
            // Retrieve it; if it does not exist, the code will simply skip duplication.
            Name originalName = null;
            try
            {
                originalName = workbook.Worksheets.Names["MyRange"];
            }
            catch { /* Named range not found – handle as needed */ }

            if (originalName != null)
            {
                // Add a new name that will be the duplicate.
                int duplicateIndex = workbook.Worksheets.Names.Add("MyRangeCopy");
                Name duplicateName = workbook.Worksheets.Names[duplicateIndex];

                // Copy the reference (RefersTo) from the original name to the duplicate.
                // The parameters indicate that the reference is in A1 style and not locale‑specific.
                duplicateName.SetRefersTo(originalName.RefersTo, false, false);
            }

            // ------------------------------------------------------------
            // Set a formula that uses the duplicated named range.
            // ------------------------------------------------------------
            // Example: place a SUM formula in cell B1 that sums the duplicated range.
            Cell targetCell = sheet.Cells["B1"];
            targetCell.Formula = "=SUM(MyRangeCopy)";

            // Calculate formulas so that the result is stored in the cell.
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // Save the modified workbook.
            // ------------------------------------------------------------
            // Replace "output.xlsx" with the desired output path.
            workbook.Save("output.xlsx");
        }
    }
}