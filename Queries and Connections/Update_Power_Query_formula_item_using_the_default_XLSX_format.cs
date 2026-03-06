using System;
using Aspose.Cells;

class UpdatePowerQueryFormulaItem
{
    static void Main()
    {
        // Load the source workbook (default XLSX format)
        Workbook workbook = new Workbook("input.xlsx");

        var dataMashup = workbook.DataMashup;
        if (dataMashup != null && dataMashup.PowerQueryFormulas != null && dataMashup.PowerQueryFormulas.Count > 0)
        {
            // Access the first Power Query formula
            var formula = dataMashup.PowerQueryFormulas[0];

            if (formula != null && formula.PowerQueryFormulaItems != null && formula.PowerQueryFormulaItems.Count > 0)
            {
                // Retrieve the first item of the formula
                var item = formula.PowerQueryFormulaItems[0];

                if (item != null && item.Value != null)
                {
                    // Example modification: replace drive letter "C:\" with "D:\"
                    string modifiedValue = item.Value.Replace(@"C:\", @"D:\");
                    item.Value = modifiedValue;

                    Console.WriteLine($"Item '{item.Name}' value updated to: {modifiedValue}");
                }
                else
                {
                    Console.WriteLine("The selected Power Query formula item is null or has no value.");
                }
            }
            else
            {
                Console.WriteLine("The selected Power Query formula contains no items.");
            }
        }
        else
        {
            Console.WriteLine("No Power Query formulas found in the workbook.");
        }

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}