// Title: Create a custom calculated item to group product categories in an Aspose.Cells pivot table (C#)
// Common Searches: how to add a calculated item to a pivot table using Aspose.Cells in C# | group multiple product categories into one pivot item with Aspose.Cells API | Aspose.Cells C# example for creating a custom pivot field item | combine pivot items with a formula in Aspose.Cells workbook | search for a pivot field by name and add a calculated item in C#
// Tags: Aspose.Cells calculated item pivot table | C# add calculated item to pivot field | group product categories Aspose.Cells | custom pivot group using formula | find pivot field by name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotField classes

// The sample loads an existing workbook, locates the first pivot table, finds the "Product Category" field, and adds a calculated item named "Food & Drinks" with the formula "Beverages+Condiments" to combine selected categories, then saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook that contains a pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Assume the first pivot table on the sheet is the target
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the field that holds product categories (e.g., "Product Category")
            string fieldName = "Product Category";

            // Name of the new calculated item that will group selected categories
            string newItemName = "Food & Drinks";

            // Formula to combine the desired categories.
            // The formula uses the names of existing items separated by '+'.
            // Adjust the item names to match those in your source data.
            string formula = "Beverages+Condiments";

            // Locate the pivot field by name (searching row, column, page, and data fields)
            PivotField targetField = FindPivotFieldByName(pivotTable, fieldName);
            if (targetField == null)
            {
                Console.WriteLine($"Error: Field \"{fieldName}\" not found in the pivot table.");
                return;
            }

            // Add the calculated item to the specified field
            targetField.AddCalculatedItem(newItemName, formula);

            // Save the workbook with the new calculated item
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to locate a pivot field by its name across all field collections
    private static PivotField FindPivotFieldByName(PivotTable pivotTable, string fieldName)
    {
        // Search row fields
        foreach (PivotField field in pivotTable.RowFields)
        {
            if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                return field;
        }

        // Search column fields
        foreach (PivotField field in pivotTable.ColumnFields)
        {
            if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                return field;
        }

        // Search page fields
        foreach (PivotField field in pivotTable.PageFields)
        {
            if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                return field;
        }

        // Search data fields (if applicable)
        foreach (PivotField field in pivotTable.DataFields)
        {
            if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                return field;
        }

        return null; // Not found
    }
}
