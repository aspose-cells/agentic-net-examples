// Title: Validate PowerQueryFormulaItem.Value Update Reflects in FormulaDefinition – Aspose.Cells .NET Example
// Description: Loads a workbook, accesses its PowerQueryFormulaCollection, modifies a PowerQueryFormulaItem.Value (e.g., swaps a drive letter), checks that the new value is present in the parent PowerQueryFormula.FormulaDefinition, and saves the workbook. Demonstrates how to confirm that item changes propagate to the query definition.
// Keywords: Aspose.Cells | PowerQueryFormulaItem | Value update | FormulaDefinition | .NET | C# | query table manipulation | workbook modification | validation | replace file path
// Common Searches: How to update PowerQueryFormulaItem.Value with Aspose.Cells | Does changing PowerQueryFormulaItem affect FormulaDefinition | Validate Power Query formula after editing items | Save workbook after Power Query changes in C# | Aspose.Cells Power Query item value replacement
// Developer Intent: Verify that setting PowerQueryFormulaItem.Value automatically updates the associated PowerQueryFormula.FormulaDefinition.
// Use Cases: Replace a hard‑coded file path in a Power Query parameter and confirm the new path appears in the formula definition before exporting. | Iterate over all PowerQueryFormulaItem objects to apply a common substring change (e.g., drive letter) and ensure each FormulaDefinition reflects the modification. | Automate a pre‑publish validation step that checks modified PowerQueryFormulaItem values are correctly embedded in their FormulaDefinition strings.
// AI Prompts: Write C# code that loops through every PowerQueryFormulaItem in a workbook, replaces a specified substring in each Value, and asserts that each FormulaDefinition contains the updated value. | Create a method that accepts a workbook path, a target string, and a replacement string, updates matching PowerQueryFormulaItem values, validates the changes in FormulaDefinition, and saves the workbook. | Explain how to log original and modified PowerQueryFormulaItem values and handle scenarios where the FormulaDefinition does not reflect the update.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // Loads a workbook, accesses its PowerQueryFormulaCollection, modifies a PowerQueryFormulaItem.Value (e.g., swaps a drive letter), checks that the new value is present in the parent PowerQueryFormula.FormulaDefinition, and saves the workbook. Demonstrates how to confirm that item changes propagate to the query definition.
    public class PowerQueryFormulaItemUpdateValidation
    {
        public static void Run()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "modified_source.xlsx";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook containing Power Query formulas
                Workbook workbook = new Workbook(sourcePath);

                // Access Power Query formulas collection
                PowerQueryFormulaCollection formulas = workbook.DataMashup.PowerQueryFormulas;

                if (formulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                    return;
                }

                // Use the first formula for demonstration
                PowerQueryFormula formula = formulas[0];

                // Access items of the selected formula
                PowerQueryFormulaItemCollection items = formula.PowerQueryFormulaItems;

                if (items.Count == 0)
                {
                    Console.WriteLine("The selected Power Query formula contains no items.");
                    return;
                }

                // Choose the first item (could be selected by name if needed)
                PowerQueryFormulaItem item = items[0];

                // Store original value for comparison
                string originalValue = item.Value;
                Console.WriteLine($"Original Item Value: {originalValue}");

                // Modify the item's value (example: replace drive letter)
                string modifiedValue = originalValue.Replace(@"C:\", @"D:\");
                item.Value = modifiedValue;
                Console.WriteLine($"Modified Item Value: {item.Value}");

                // Verify that the change is reflected in the formula definition
                string updatedDefinition = formula.FormulaDefinition;
                Console.WriteLine($"Updated Formula Definition: {updatedDefinition}");

                bool isChangeReflected = updatedDefinition.Contains(modifiedValue);
                Console.WriteLine(isChangeReflected
                    ? "The change in PowerQueryFormulaItem.Value is reflected in the formula definition."
                    : "The change was NOT reflected in the formula definition.");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PowerQueryFormulaItemUpdateValidation.Run();
        }
    }
}
