// Title: Rename an Excel worksheet and recalculate formulas that reference it using Aspose.Cells for .NET (C#)
// AI Prompts: Rename a worksheet in a Workbook, then invoke Workbook.CalculateFormula to update all formulas that still point to the original sheet name. | Create a workbook, assign a formula that includes the sheet name, change the sheet's name, and refresh the formula result with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# how to update formulas after renaming a worksheet | recalculate Excel formulas when sheet name changes using Aspose.Cells .NET | Workbook.CalculateFormula usage after worksheet rename in C# | refresh formula results that reference old sheet name with Aspose.Cells
// Tags: worksheet rename refresh formulas Aspose.Cells | Workbook.CalculateFormula after sheet rename | update formula references C# Aspose.Cells | Excel sheet name change recalculate formulas .NET | Aspose.Cells rename worksheet and recalc

using System;
using Aspose.Cells;

namespace AsposeCellsRenameAndRefresh
{
    // The example creates a new workbook, adds values and a formula that explicitly references the worksheet name, renames the worksheet, calls Workbook.CalculateFormula to refresh the formula, outputs the result, and saves the file as RenamedAndRefreshed.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some data and a formula that references the sheet name
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            // Formula uses the sheet name explicitly (e.g., Sheet1!A1+A2)
            sheet.Cells["B1"].Formula = $"={sheet.Name}!A1+{sheet.Name}!A2";

            // Rename the worksheet
            string oldName = sheet.Name;
            string newName = "RenamedSheet";
            sheet.Name = newName;

            // After renaming, recalculate all formulas in the workbook
            // This will refresh formulas that referenced the old sheet name
            workbook.CalculateFormula();

            // Output the result of the formula to verify it was refreshed
            Console.WriteLine($"Old sheet name: {oldName}");
            Console.WriteLine($"New sheet name: {newName}");
            Console.WriteLine($"Formula result in B1: {sheet.Cells["B1"].Value}");

            // Save the workbook (lifecycle save) – optional for verification
            workbook.Save("RenamedAndRefreshed.xlsx");
        }
    }
}
