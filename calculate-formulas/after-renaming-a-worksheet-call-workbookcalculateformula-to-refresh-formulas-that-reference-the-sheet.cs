// Title: Rename a worksheet and refresh dependent formulas using Workbook.CalculateFormula in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to rename a worksheet, populate data, and call Workbook.CalculateFormula to automatically update all formulas that reference the renamed sheet, then retrieve the recalculated result.
// Keywords: Aspose.Cells rename worksheet | Workbook.CalculateFormula C# | update formula references after sheet rename | refresh formulas Aspose.Cells .NET | programmatic sheet rename Aspose.Cells | calculate all formulas Aspose.Cells
// Common Searches: Aspose.Cells update formulas after renaming sheet | Workbook.CalculateFormula example C# | how to refresh formulas when sheet name changes Aspose.Cells | C# rename worksheet and recalculate formulas | Aspose.Cells formula refresh after sheet rename
// Developer Intent: Recalculate all workbook formulas so they reflect a worksheet’s new name.
// Use Cases: Rename a data sheet in a generated report and keep summary calculations accurate. | Automate batch renaming of worksheets while preserving dependent formulas. | Create dynamic workbooks where sheet names change based on user input and formulas must stay valid.
// AI Prompts: Show a C# example that renames a worksheet and then calls Workbook.CalculateFormula to update formulas in Aspose.Cells. | Explain why Workbook.CalculateFormula is required after changing a sheet name and what alternatives exist. | Provide step‑by‑step code to rename a sheet, refresh formulas, and read the calculated value using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRenameAndRefresh
{
    // Demonstrates how to rename a worksheet, populate data, and call Workbook.CalculateFormula to automatically update all formulas that reference the renamed sheet, then retrieve the recalculated result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the default first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "OriginalSheet";

            // Add a second worksheet that contains a formula referencing the first sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Summary");
            // Example: sum of A1:A3 from the first sheet
            sheet2.Cells["A1"].Formula = "=SUM(OriginalSheet!A1:A3)";

            // Populate some data in the first sheet
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            // Rename the first worksheet
            sheet1.Name = "RenamedSheet";

            // Refresh all formulas in the workbook so that references to the renamed sheet are updated
            workbook.CalculateFormula();

            // Output the result of the formula after refresh
            Console.WriteLine("Result of formula in Summary!A1 after renaming: " + sheet2.Cells["A1"].Value);
        }
    }
}
