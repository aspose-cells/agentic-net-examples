// Title: Create a dynamic array formula referencing a table column and capture the #REF! error after deleting the table with Aspose.Cells for .NET
// Common Searches: how to insert a SORT dynamic array formula that uses a ListObject column in Aspose.Cells C# | Aspose.Cells C# retrieve #REF! error when a structured table reference is removed | recalculate workbook formulas after deleting a ListObject in Aspose.Cells | example of using structured references with dynamic arrays in Aspose.Cells .NET | save workbook after dynamic array formula spills and table deletion using Aspose.Cells
// Tags: dynamic array formula with structured table reference | Aspose.Cells SORT function on ListObject column | handle #REF! error after table removal | recalculate formulas after ListObject deletion | C# Aspose.Cells create and delete ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample builds a workbook, fills column A with fruit names, creates a ListObject named MyTable with a Fruit column, inserts a =SORT(MyTable[Fruit]) dynamic array formula in B1, calculates and prints the spilled results, then removes the table, recalculates, verifies that the formula cell returns a #REF! error, and saves the workbook.
class DynamicArrayFormulaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data in column A (A1:A5)
            string[] data = { "Apple", "Banana", "Apple", "Cherry", "Banana" };
            for (int i = 0; i < data.Length; i++)
            {
                cells[i, 0].PutValue(data[i]); // A1, A2, ...
            }

            // Create a table (ListObject) over the data range A1:A5
            int firstRow = 0, firstCol = 0, totalRows = data.Length, totalCols = 1;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, firstRow + totalRows - 1, firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set table name (use DisplayName for compatibility)
            table.DisplayName = "MyTable";

            // Set a column name for the table (required for structured references)
            table.ListColumns[0].Name = "Fruit";

            // Insert a dynamic array formula that references the table column.
            // The formula will spill into adjacent cells starting at B1.
            cells[0, 1].Formula = "=SORT(MyTable[Fruit])";

            // Calculate all formulas
            workbook.CalculateFormula();

            // Display the results of the array formula before deleting the table
            Console.WriteLine("Array formula results before deleting the table:");
            for (int r = 0; r < data.Length; r++)
            {
                // The spilled results start at B1 and occupy B1:B5
                Console.WriteLine(cells[r, 1].StringValue);
            }

            // Delete the table (ListObject) from the worksheet
            sheet.ListObjects.RemoveAt(0); // Removes "MyTable"

            // Recalculate formulas after the table has been removed
            workbook.CalculateFormula();

            // Observe the formula error after the table deletion
            Console.WriteLine("\nAfter deleting the table, the formula cell contains:");
            Cell formulaCell = cells[0, 1];
            if (formulaCell.Type == CellValueType.IsError)
            {
                // Expected: #REF! error because the structured reference is no longer valid
                Console.WriteLine(formulaCell.StringValue); // Should display "#REF!"
            }
            else
            {
                // If not an error, display whatever value is present
                Console.WriteLine(formulaCell.StringValue);
            }

            // Optionally, save the workbook to inspect the result manually
            string outputPath = "DynamicArrayFormulaDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
