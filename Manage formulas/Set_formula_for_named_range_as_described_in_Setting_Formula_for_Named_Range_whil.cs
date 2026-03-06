using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeFormula
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook without parsing formulas on open.
            // This allows us to modify the named range formula before any calculation occurs.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false; // skip formula parsing during load
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Ensure the workbook has at least one worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1"; // make sure the sheet name matches the formula reference

            // Add a new named range (or retrieve an existing one) and set its formula.
            // The RefersTo property defines the range the name points to, beginning with '='.
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$5"; // set the formula for the named range

            // Optionally, calculate all formulas now that the named range is defined.
            workbook.CalculateFormula();

            // Save the modified workbook.
            workbook.Save("output.xlsx");
        }
    }
}