using System;
using Aspose.Cells;

class UpdatePowerQueryFormulaItem
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the DataMashup object
        var dataMashup = workbook.DataMashup;

        if (dataMashup != null && dataMashup.PowerQueryFormulas != null && dataMashup.PowerQueryFormulas.Count > 0)
        {
            // Get the first Power Query formula
            var formula = dataMashup.PowerQueryFormulas[0];

            if (formula != null && formula.PowerQueryFormulaItems != null && formula.PowerQueryFormulaItems.Count > 0)
            {
                // Get the first item
                var item = formula.PowerQueryFormulaItems[0];

                // Example modification: replace drive letter "C:\" with "D:\"
                string modifiedValue = item.Value.Replace(@"C:\", @"D:\");

                // Update the item's value
                item.Value = modifiedValue;

                // Output the change for verification
                Console.WriteLine("Power Query formula item value updated to: " + modifiedValue);
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

        // Save the modified workbook (default XLSX format)
        workbook.Save("output.xlsx");
    }
}