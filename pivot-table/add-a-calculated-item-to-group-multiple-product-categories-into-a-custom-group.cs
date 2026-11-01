// Title: Create a custom calculated item to combine product categories in an Excel PivotTable with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add calculated item to pivot table field | group product categories into a single item in Excel pivot using Aspose.Cells | how to define a calculated item formula for a pivot field in Aspose.Cells .NET | C# code to sum specific categories in a pivot table with Aspose.Cells | add custom group to pivot table rows using Aspose.Cells API
// Tags: calculated item pivotfield Aspose.Cells | product category custom group pivot Aspose.Cells | pivot table calculated item formula C# | Aspose.Cells Excel pivot manipulation | custom group pivot rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace Example
{
    // The example loads an Excel workbook, locates the first pivot table, finds the 'Product Category' field, and adds a calculated item named 'Custom Group' with the formula ='Category A' + 'Category B' to group those categories, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load workbook (no password)
                Workbook workbook = new Workbook(inputPath);

                // Access first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Name of the pivot field that contains product categories
                string fieldName = "Product Category";

                // Locate the field in row or column fields
                PivotField productCategoryField = FindPivotField(pivotTable, fieldName);

                if (productCategoryField != null)
                {
                    // Define the name of the new calculated item
                    string calculatedItemName = "Custom Group";

                    // Define the formula that sums the desired categories.
                    // The formula syntax uses single quotes around item names.
                    string formula = "='Category A' + 'Category B'";

                    // Add the calculated item to the pivot field
                    productCategoryField.AddCalculatedItem(calculatedItemName, formula);
                }
                else
                {
                    Console.WriteLine($"Pivot field \"{fieldName}\" not found.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper method to locate a pivot field by name
        private static PivotField FindPivotField(PivotTable pivotTable, string fieldName)
        {
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return field;
            }
            foreach (PivotField field in pivotTable.ColumnFields)
            {
                if (field.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return field;
            }
            return null;
        }
    }
}
