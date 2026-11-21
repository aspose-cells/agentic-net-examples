// Title: Limit shared formula rows to 50 with Workbook.Settings.MaxRowsOfSharedFormula in Aspose.Cells C#
// Common Searches: Aspose.Cells C# limit shared formula rows to 50 | set MaxRowsOfSharedFormula property in .NET workbook | how to restrict number of rows using shared formulas in Aspose.Cells | configure workbook settings to cap shared formula rows in C# | example of limiting shared formula rows with Aspose.Cells API
// Tags: maxrowsofsharedformula c# | aspocells workbook shared formula limit | c# workbook settings max rows shared formula | limit shared formula rows aspocells | set shared formula row cap .net

using Aspose.Cells;
using System;

// The sample creates a new Workbook, sets the Settings.MaxRowsOfSharedFormula property to 50 to restrict how many rows can share a formula, and saves the file as SharedFormulaLimited.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Limit shared formula rows to fifty
        workbook.Settings.MaxRowsOfSharedFormula = 50;

        // Save the workbook
        workbook.Save("SharedFormulaLimited.xlsx");
    }
}
