// Title: Generate a UNIQUE dynamic array formula referencing a ListObject column, delete the table, and detect the resulting #REF! error with Aspose.Cells for .NET
// AI Prompts: Create a workbook, add a ListObject named MyTable over A1:A5, set B1 formula to =UNIQUE(MyTable[Fruit]), calculate, then remove the ListObject and return the error value from B1. | Using Aspose.Cells in C#, insert a dynamic array formula that extracts distinct values from a table column, delete the table, recalculate the workbook, and output the #REF! error produced.
// Common Searches: Aspose.Cells how to get #REF error after deleting a table used in a UNIQUE formula | C# dynamic array formula referencing ListObject column then removing the table | calculate formulas after removing ListObject in Aspose.Cells .NET | retrieve distinct values from a table column with UNIQUE function in Aspose.Cells
// Tags: dynamic array formula UNIQUE Aspose.Cells | ListObject column reference formula C# | detect #REF error after table deletion Aspose.Cells | calculate workbook after ListObject removal .NET | distinct values from spreadsheet table column using UNIQUE

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a ListObject named MyTable over cells A1:A5, places a UNIQUE dynamic array formula =UNIQUE(MyTable[Fruit]) in B1, calculates and prints the distinct fruit values, then removes the table, recalculates, and displays the #REF! error that appears in B1 before saving the file.
class DynamicArrayFormulaDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the table in column A (A1:A5)
            string[] data = { "Apple", "Banana", "Apple", "Cherry", "Banana" };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(data[i]); // Column A
            }

            // Create a table (ListObject) covering A1:A5 and name it "MyTable"
            int firstRow = 0, firstCol = 0, totalRows = data.Length, totalCols = 1;
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, firstRow + totalRows, firstCol + totalCols, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";
            table.ShowHeaderRow = true; // First row as header

            // Set header name for the column
            sheet.Cells[firstRow, firstCol].PutValue("Fruit");

            // Insert a dynamic array formula in B1 that references the table column
            // Using UNIQUE to return distinct values from the table column
            Cell formulaCell = sheet.Cells[0, 1]; // Cell B1
            formulaCell.Formula = "=UNIQUE(MyTable[Fruit])";

            // Calculate formulas to populate the result
            workbook.CalculateFormula();

            // Output the results before deleting the table
            Console.WriteLine("Results before deleting the table:");
            int resultRow = 0;
            while (!string.IsNullOrEmpty(sheet.Cells[resultRow, 1].StringValue))
            {
                Console.WriteLine($"B{resultRow + 1}: {sheet.Cells[resultRow, 1].StringValue}");
                resultRow++;
            }

            // Delete the table (ListObject)
            sheet.ListObjects.RemoveAt(0);

            // Recalculate formulas after table deletion
            workbook.CalculateFormula();

            // Check the formula cell for error after deletion
            Cell errorCell = sheet.Cells[0, 1]; // B1
            string errorValue = errorCell.StringValue;
            if (errorValue.StartsWith("#"))
            {
                Console.WriteLine("\nAfter deleting the table, the formula results in an error:");
                Console.WriteLine($"Error in B1: {errorValue}"); // Expected #REF!
            }
            else
            {
                Console.WriteLine("\nAfter deleting the table, the formula did not produce an error (unexpected).");
                Console.WriteLine($"B1 value: {errorValue}");
            }

            // Save the workbook to a file (optional)
            workbook.Save("DynamicArrayFormulaDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
