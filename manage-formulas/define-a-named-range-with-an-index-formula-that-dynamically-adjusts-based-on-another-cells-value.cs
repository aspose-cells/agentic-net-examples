// Title: Define a dynamic named range with INDEX that reacts to a cell value – Aspose.Cells for .NET example
// Description: Shows how to create a workbook, populate A1:A10, place an index in B1, add a named range "DynamicItem" using =INDEX(Sheet1!$A$1:$A$10, Sheet1!$B$1), reference the name in C1, calculate formulas, retrieve the resolved address, and save the file.
// Keywords: Aspose.Cells | .NET | C# | dynamic named range | INDEX function | named range formula | retrieve named range address | calculate formulas | Excel automation
// Common Searches: Aspose.Cells create named range with INDEX | dynamic named range based on another cell .NET | how to use INDEX in a named range with Aspose.Cells | retrieve resolved address of a named range Aspose.Cells C# | update named range when index cell changes
// Developer Intent: Create a named range whose reference is defined by an INDEX formula that uses the value of another cell to select the target row.
// Use Cases: Show a specific list item in a summary cell by changing the index value. | Use the same dynamically selected item in multiple calculations without hard‑coding the address. | Build parameter‑driven reports where a user‑editable cell determines which row of data is displayed.
// AI Prompts: Modify the example so the index cell (B1) is itself a named range and the INDEX formula references that name. | Extend the code to use a two‑dimensional INDEX (row, column) to return a sub‑range from a table and assign it to a named range. | Explain how to keep the named range formula valid when the worksheet is renamed or moved.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicNamedRange
{
    // Shows how to create a workbook, populate A1:A10, place an index in B1, add a named range "DynamicItem" using =INDEX(Sheet1!$A$1:$A$10, Sheet1!$B$1), reference the name in C1, calculate formulas, retrieve the resolved address, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a vertical list of values in column A (A1:A10)
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Cell B1 will hold the index (row number) to pick from the list
                // For demonstration set it to 3 (will pick "Item 3")
                cells[0, 1].PutValue(3);

                // Add a named range "DynamicItem" that uses INDEX to refer to a single cell
                // Formula: =INDEX(Sheet1!$A$1:$A$10, Sheet1!$B$1)
                int nameIndex = workbook.Worksheets.Names.Add("DynamicItem");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = $"=INDEX({sheet.Name}!$A$1:$A$10, {sheet.Name}!$B$1)";

                // Use the named range in a formula (e.g., cell C1)
                cells["C1"].Formula = "=DynamicItem";

                // Calculate all formulas so that C1 gets the value from the indexed cell
                workbook.CalculateFormula();

                // Retrieve the range that the name currently refers to (after calculation)
                AsposeRange resolvedRange = dynamicName.GetRange(true);
                Console.WriteLine($"DynamicItem resolves to address: {resolvedRange.RefersTo}");
                Console.WriteLine($"Value in C1 (should match indexed item): {cells["C1"].StringValue}");

                // Save the workbook to verify the named range and result
                string outputPath = "DynamicNamedRangeDemo.xlsx";
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
