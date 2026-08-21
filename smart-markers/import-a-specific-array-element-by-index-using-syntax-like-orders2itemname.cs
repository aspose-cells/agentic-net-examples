// Title: Import a Specific Array Element by Index into Excel with Aspose.Cells for .NET
// Description: Demonstrates how to place a single array item into a worksheet cell using Aspose.Cells. The example shows setting a Power Query formula that references an element by index (e.g., =Orders[2].ItemName), retrieving a PowerQueryFormulaItem at a given position, and importing the third value of a string array into cell C1 with ImportObjectArray.
// Keywords: Aspose.Cells ImportObjectArray | C# array element index Excel | Power Query formula index Aspose | smart markers array access | .NET Excel cell value from array | Aspose.Cells example
// Common Searches: Aspose.Cells import single array value | How to reference Power Query table row by index in formula | C# ImportObjectArray specific element | Set Excel cell to array item using Aspose | Access PowerQueryFormulaItem by index
// Developer Intent: Write a single array element to an Excel cell and work with indexed Power Query items using Aspose.Cells.
// Use Cases: Insert the third entry of a C# string array into a designated cell without loading the whole array. | Create a cell formula that points to a Power Query table row by zero‑based index (e.g., =Orders[2].ItemName). | Retrieve a PowerQueryFormulaItem at a specific position, modify its value, and write the result back to the worksheet.
// AI Prompts: Generate C# code that uses Aspose.Cells ImportObjectArray to write the fourth element of an integer array into cell D5. | Show how to update the value of a PowerQueryFormulaItem at index 5 and output the new value to cell A1. | Provide an example of setting a cell formula that references the second row of a Power Query table named "Sales" using the &=Sales[1].ColumnName syntax.

using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsExamples
{
    // Demonstrates how to place a single array item into a worksheet cell using Aspose.Cells. The example shows setting a Power Query formula that references an element by index (e.g., =Orders[2].ItemName), retrieving a PowerQueryFormulaItem at a given position, and importing the third value of a string array into cell C1 with ImportObjectArray.
    public class ImportArrayElementByIndexDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Example 1: Set a cell formula that references a Power Query
            // table element by index (syntax similar to &=Orders[2].ItemName)
            // ------------------------------------------------------------
            // Note: The actual Power Query table "Orders" must exist in the workbook.
            // For demonstration, we simply assign the formula string.
            // In a real scenario, load a workbook that contains the Power Query table.
            cells["A1"].Formula = "=Orders[2].ItemName";

            // ------------------------------------------------------------
            // Example 2: Access a Power Query formula item by numeric index
            // using the PowerQueryFormulaItemCollection indexer.
            // ------------------------------------------------------------
            // Ensure that the workbook contains at least one Power Query formula.
            // If not, this block will be skipped.
            PowerQueryFormulaCollection pqFormulas = workbook.DataMashup.PowerQueryFormulas;
            if (pqFormulas != null && pqFormulas.Count > 0)
            {
                // Get the first Power Query formula
                PowerQueryFormula formula = pqFormulas[0];

                // Access its collection of formula items
                PowerQueryFormulaItemCollection items = formula.PowerQueryFormulaItems;

                // Verify that the collection has enough items
                if (items != null && items.Count > 2) // we want the item at index 2 (third item)
                {
                    // Retrieve the item at index 2
                    PowerQueryFormulaItem thirdItem = items[2];

                    // For demonstration, set a new value for this item
                    thirdItem.Value = "NewValueForThirdItem";

                    // Optionally, write the value back to a cell for verification
                    cells["B1"].PutValue(thirdItem.Value);
                }
            }

            // ------------------------------------------------------------
            // Example 3: Import a single string element from an array into a cell
            // ------------------------------------------------------------
            string[] sampleArray = new string[] { "Alpha", "Beta", "Gamma", "Delta" };
            // Import only the element at index 2 ("Gamma") into cell C1
            // Using ImportObjectArray with a single-element array and vertical=false
            sheet.Cells.ImportObjectArray(new object[] { sampleArray[2] }, 0, 2, false);

            // ------------------------------------------------------------
            // Save the workbook (lifecycle rule: save)
            // ------------------------------------------------------------
            workbook.Save("ImportArrayElementByIndexDemo.xlsx");
        }
    }
}
