// Title: C# Example: Update and Verify PowerQueryFormulaItem.Value with Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel workbook, access its DataMashup part, modify the Value of the first PowerQueryFormulaItem (e.g., change a file‑path), save the workbook, reload it, and confirm that the change persists. This validates that PowerQueryFormulaItem updates are correctly written to the file using Aspose.Cells.
// Keywords: Aspose.Cells | PowerQueryFormulaItem | Power Query update .NET | C# Excel DataMashup | modify Power Query formula value | persist Power Query changes | validate PowerQueryFormulaItem | Excel workbook automation | Aspose.Cells example
// Common Searches: how to change PowerQueryFormulaItem value with Aspose.Cells | verify Power Query formula updates are saved in Excel C# | Aspose.Cells modify PowerQueryFormulaItem and reload workbook | C# code to edit Power Query source definition using Aspose | Aspose.Cells DataMashup PowerQueryFormulas example
// Developer Intent: Edit a PowerQueryFormulaItem.Value in an Excel file and ensure the modification is saved and retrievable.
// Use Cases: Replace environment‑specific file paths in Power Query source definitions before distributing workbooks. | Automate regression tests that edit Power Query formulas and verify persistence after a save/load cycle. | Batch‑process multiple workbooks to adjust PowerQueryFormulaItems for different deployment configurations.
// AI Prompts: Generate C# code that iterates through all PowerQueryFormulaItems in a workbook and replaces a given substring in each Value using Aspose.Cells. | Provide a step‑by‑step guide to modify a PowerQueryFormulaItem.Value, save the workbook, reload it, and assert the new value matches the expected result. | Explain how to handle missing DataMashup parts or empty PowerQueryFormulas collections when updating Power Query formulas with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an Excel workbook, access its DataMashup part, modify the Value of the first PowerQueryFormulaItem (e.g., change a file‑path), save the workbook, reload it, and confirm that the change persists. This validates that PowerQueryFormulaItem updates are correctly written to the file using Aspose.Cells.
class ValidatePowerQueryFormulaItemUpdate
{
    static void Main()
    {
        try
        {
            // Input workbook containing Power Query formulas
            string inputPath = "PowerQuerySource.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access DataMashup part via dynamic to avoid compile‑time dependency on Aspose.Cells.DataMashup assembly
            dynamic dataMashup = workbook.DataMashup;
            if (dataMashup == null)
            {
                Console.WriteLine("The workbook does not contain a DataMashup part.");
                return;
            }

            // Verify that at least one Power Query formula exists
            var powerQueryFormulas = dataMashup.PowerQueryFormulas;
            if (powerQueryFormulas == null || powerQueryFormulas.Count == 0)
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
                return;
            }

            // Access the first Power Query formula
            dynamic formula = powerQueryFormulas[0];

            // Verify that the formula has at least one item
            var items = formula.PowerQueryFormulaItems;
            if (items == null || items.Count == 0)
            {
                Console.WriteLine("The selected Power Query formula contains no items.");
                return;
            }

            // Retrieve the first item
            dynamic item = items[0];
            Console.WriteLine($"Original Item Name : {item.Name}");
            Console.WriteLine($"Original Item Value: {item.Value}");

            // Modify the item's Value (example: replace drive letter C:\ with D:\)
            string modifiedValue = ((string)item.Value).Replace(@"C:\", @"D:\");
            item.Value = modifiedValue;
            Console.WriteLine($"Modified Item Value: {item.Value}");

            // Save the workbook with the updated value
            string outputPath = "PowerQueryModified.xlsx";
            workbook.Save(outputPath);

            // Reload the saved workbook to verify that the change persisted
            Workbook reloadedWorkbook = new Workbook(outputPath);
            dynamic reloadedDataMashup = reloadedWorkbook.DataMashup;
            if (reloadedDataMashup == null)
            {
                Console.WriteLine("Reloaded workbook does not contain a DataMashup part.");
                return;
            }

            dynamic reloadedFormula = reloadedDataMashup.PowerQueryFormulas[0];
            dynamic reloadedItem = reloadedFormula.PowerQueryFormulaItems[0];
            Console.WriteLine($"Reloaded Item Value: {reloadedItem.Value}");

            // Validate that the reloaded value matches the modified value
            bool validationPassed = reloadedItem.Value == modifiedValue;
            Console.WriteLine($"Validation result: {(validationPassed ? "Success" : "Failure")}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
