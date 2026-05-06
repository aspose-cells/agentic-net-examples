using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryUpdate
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains Power Query formulas
            Workbook workbook = new Workbook("source.xlsx");

            // Ensure the workbook has DataMashup and at least one Power Query formula
            DataMashup mashup = workbook.DataMashup;
            if (mashup == null || mashup.PowerQueryFormulas.Count == 0)
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
                return;
            }

            // Access the first Power Query formula
            PowerQueryFormula formula = mashup.PowerQueryFormulas[0];

            // Ensure the formula has at least one item
            if (formula.PowerQueryFormulaItems.Count == 0)
            {
                Console.WriteLine("The selected Power Query formula contains no items.");
                return;
            }

            // Access the first item of the formula
            PowerQueryFormulaItem item = formula.PowerQueryFormulaItems[0];

            // Display original value
            Console.WriteLine("Original Item Value: " + item.Value);

            // Example modification: replace a drive letter in the item value
            string modifiedValue = item.Value.Replace(@"C:\", @"D:\");
            item.Value = modifiedValue;

            // Verify the change
            Console.WriteLine("Modified Item Value: " + item.Value);

            // Save the workbook in the default XLSX format
            workbook.Save("modified_source.xlsx");
            Console.WriteLine("Workbook saved as modified_source.xlsx");
        }
    }
}