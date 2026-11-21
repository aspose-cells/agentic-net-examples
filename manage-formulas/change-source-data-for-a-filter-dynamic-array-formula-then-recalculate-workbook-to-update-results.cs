// Title: Updating the source range of a FILTER dynamic array formula and recalculating the workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace FILTER formula source range | Change array argument of FILTER function in Excel file using Aspose.Cells | Recalculate workbook after modifying dynamic array formula with Aspose.Cells .NET | Update FILTER formula range programmatically in .xlsx using Aspose.Cells | Use Regex to edit Excel formula in Aspose.Cells C#
// Tags: modify FILTER function array argument Aspose.Cells | recalculate workbook formulas Aspose.Cells .NET | dynamic array formula update Aspose.Cells | regex-based formula editing Aspose.Cells | assign updated formula to cell Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an existing workbook, locates a cell containing a FILTER dynamic array formula, replaces the first argument (the source range) using a regular expression, writes the updated formula back to the cell, forces a full workbook recalculation to refresh the dynamic array results, and saves the modified file.
class FilterFormulaUpdater
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the cell that contains the FILTER dynamic array formula.
        // For this example we assume it is in cell A1.
        // If the location is unknown, you could iterate through cells and check for "FILTER(" in the formula.
        Cell filterCell = sheet.Cells["A1"];

        // Ensure the cell actually contains a formula
        if (filterCell.IsFormula && filterCell.Formula.Contains("FILTER("))
        {
            // Original formula, e.g., =FILTER(Sheet1!$B$2:$B$10, Sheet1!$C$2:$C$10)
            string originalFormula = filterCell.Formula;

            // Define the new source range you want the FILTER function to use.
            // Example: change the array argument to Sheet1!$D$2:$D$20
            string newArrayRange = "Sheet1!$D$2:$D$20";

            // Use a regular expression to replace the first argument of FILTER (the array range)
            // FILTER\(\s*([^,]+) captures the first argument up to the first comma.
            string pattern = @"FILTER\(\s*([^,]+)";
            string replacement = $"FILTER({newArrayRange}";
            string updatedFormula = Regex.Replace(originalFormula, pattern, replacement, RegexOptions.IgnoreCase);

            // Apply the updated formula back to the cell
            filterCell.Formula = updatedFormula;
        }
        else
        {
            Console.WriteLine("The specified cell does not contain a FILTER formula.");
            return;
        }

        // Recalculate the workbook so that the dynamic array results are refreshed
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
