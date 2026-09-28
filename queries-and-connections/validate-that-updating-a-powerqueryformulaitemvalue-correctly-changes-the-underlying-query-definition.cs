// Title: Validate that changing a PowerQueryFormulaItem.Value updates the Power Query formula definition in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, retrieves the first PowerQueryFormulaItem, modifies its Value, compares the original FormulaDefinition with the updated one, and saves the workbook. | Show how to programmatically replace a file path inside a PowerQueryFormulaItem.Value and verify that the workbook's DataMashup reflects the change. | Provide a C# snippet that logs the original and new values of a PowerQueryFormulaItem, checks whether FormulaDefinition has changed, and outputs a success message.
// Common Searches: Aspose.Cells C# update PowerQueryFormulaItem value and confirm formula definition alteration | How to edit a Power Query formula item in an Excel workbook using Aspose.Cells .NET library | Example code to check if DataMashup FormulaDefinition updates after changing a Power Query item | C# verify Power Query formula definition change after modifying item value with Aspose.Cells
// Tags: Aspose.Cells modify PowerQueryFormulaItem value | Aspose.Cells verify DataMashup formula definition change | C# update Power Query item in Excel workbook | Save workbook after Power Query item edit with Aspose.Cells | PowerQueryFormulaItem value replacement C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // // Loads an Excel workbook, accesses its DataMashup, updates the first PowerQueryFormulaItem's Value, compares the original and updated FormulaDefinition to confirm the change, and saves the modified workbook.
    public class PowerQueryFormulaItemValueUpdateValidation
    {
        // Entry point for the application.
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                string sourcePath = "source.xlsx";

                // Verify that the source workbook exists.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the workbook containing Power Query formulas.
                Workbook workbook = new Workbook(sourcePath);

                // Access the DataMashup object which holds Power Query formulas.
                DataMashup mashup = workbook.DataMashup;

                // Ensure there is at least one Power Query formula.
                if (mashup == null || mashup.PowerQueryFormulas.Count == 0)
                {
                    Console.WriteLine("No Power Query formulas found in the workbook.");
                    return;
                }

                // Get the first Power Query formula.
                PowerQueryFormula formula = mashup.PowerQueryFormulas[0];

                // Ensure the formula contains at least one item.
                if (formula.PowerQueryFormulaItems.Count == 0)
                {
                    Console.WriteLine("The selected Power Query formula contains no items.");
                    return;
                }

                // Capture the original formula definition for later comparison.
                string originalDefinition = formula.FormulaDefinition;

                // Access the first item of the formula.
                PowerQueryFormulaItem item = formula.PowerQueryFormulaItems[0];

                // Display original item value.
                Console.WriteLine($"Original Item Name : {item.Name}");
                Console.WriteLine($"Original Item Value: {item.Value}");

                // Modify the item's value.
                string modifiedValue = item.Value.Contains(@"C:\")
                    ? item.Value.Replace(@"C:\", @"D:\")
                    : item.Value + "_Modified";

                item.Value = modifiedValue;

                // Display the modified item value.
                Console.WriteLine($"Modified Item Value: {item.Value}");

                // Retrieve the updated formula definition.
                string updatedDefinition = formula.FormulaDefinition;

                // Validate that the formula definition has changed.
                bool definitionChanged = !string.Equals(originalDefinition, updatedDefinition, StringComparison.Ordinal);
                Console.WriteLine($"Formula definition changed: {definitionChanged}");

                if (definitionChanged)
                {
                    Console.WriteLine("Update successful. New formula definition:");
                    Console.WriteLine(updatedDefinition);
                }
                else
                {
                    Console.WriteLine("Update failed. Formula definition remains unchanged.");
                }

                // Save the workbook with the modifications.
                string outputPath = "modified_source.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
