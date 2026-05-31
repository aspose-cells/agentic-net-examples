using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryValidation
{
    class Program
    {
        static void Main()
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "modified_output.xlsx";

            try
            {
                // Verify source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: '{sourcePath}'.");
                    return;
                }

                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(sourcePath);

                // Access the DataMashup which holds Power Query formulas
                DataMashup mashup = workbook.DataMashup;
                if (mashup == null)
                {
                    Console.WriteLine("The workbook does not contain any Power Query data mashup.");
                    return;
                }

                // Ensure there is at least one Power Query formula
                if (mashup.PowerQueryFormulas == null || mashup.PowerQueryFormulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                    return;
                }

                // Get the first Power Query formula
                PowerQueryFormula formula = mashup.PowerQueryFormulas[0];
                if (formula == null)
                {
                    Console.WriteLine("Failed to retrieve the Power Query formula.");
                    return;
                }

                // Capture the original formula definition for later comparison
                string originalDefinition = formula.FormulaDefinition;

                // Ensure the formula has at least one item
                if (formula.PowerQueryFormulaItems == null || formula.PowerQueryFormulaItems.Count == 0)
                {
                    Console.WriteLine("The selected Power Query formula contains no items.");
                    return;
                }

                // Access the first item (could also use name‑based indexer)
                PowerQueryFormulaItem item = formula.PowerQueryFormulaItems[0];
                if (item == null)
                {
                    Console.WriteLine("Failed to retrieve the Power Query formula item.");
                    return;
                }

                // Display original item value
                Console.WriteLine($"Original Item Name : {item.Name}");
                Console.WriteLine($"Original Item Value: {item.Value}");

                // Modify the item's value – for demonstration replace a drive letter if present
                string modifiedValue = item.Value?.Replace(@"C:\", @"D:\") ?? string.Empty;

                // Assign the new value back to the item
                item.Value = modifiedValue;

                // Display the modified item value
                Console.WriteLine($"Modified Item Value: {item.Value}");

                // After modification, retrieve the updated formula definition
                string updatedDefinition = formula.FormulaDefinition;

                // Validate that the definition has changed
                bool definitionChanged = !string.Equals(originalDefinition, updatedDefinition, StringComparison.Ordinal);
                Console.WriteLine($"Formula definition changed: {definitionChanged}");

                if (definitionChanged)
                {
                    Console.WriteLine("Original Definition:");
                    Console.WriteLine(originalDefinition);
                    Console.WriteLine("Updated Definition:");
                    Console.WriteLine(updatedDefinition);
                }
                else
                {
                    Console.WriteLine("Formula definition did not change – update may have failed.");
                }

                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}