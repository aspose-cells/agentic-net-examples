// Title: Aspose.Cells .NET: Capture Expected Exception When Filtering a Table After ConvertToRange
// Description: Demonstrates how to create a workbook with a ListObject (table) that has an auto‑filter, convert the table to a normal range, trigger and catch the exception thrown by ListObject.Filter, then apply a worksheet‑level filter to the same cell area and save the file.
// Keywords: Aspose.Cells | .NET | C# | ListObject | ConvertToRange | Filter exception | auto‑filter | worksheet filter | range filtering | sample code | demo
// Common Searches: Aspose.Cells ListObject.Filter after ConvertToRange exception | how to catch exception when filtering a converted table in C# | apply worksheet filter to range formerly a table Aspose.Cells | ConvertToRange removes table auto‑filter Aspose.Cells | sample code for table to range conversion Aspose.Cells
// Developer Intent: The developer wants to confirm that calling ListObject.Filter after a table is converted to a range throws an exception, capture that exception, and then use a worksheet‑level filter on the same cells.
// Use Cases: Validate that ConvertToRange disables the ListObject.Filter method. | Show how to handle the expected exception gracefully. | Apply a worksheet filter to a range that was previously a table. | Save the workbook after exception handling and filter application.
// AI Prompts: Generate C# code with Aspose.Cells that creates a table, converts it to a range, attempts ListObject.Filter, catches the expected exception, and then applies a worksheet filter to the same area. | Explain why ListObject.Filter throws an exception after ConvertToRange and describe the correct way to filter the range at the worksheet level. | Write a unit test in C# that asserts an exception is thrown when invoking ListObject.Filter on a converted range and verifies that a worksheet filter can still be applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with a ListObject (table) that has an auto‑filter, convert the table to a normal range, trigger and catch the exception thrown by ListObject.Filter, then apply a worksheet‑level filter to the same cell area and save the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                FilterAfterTableConversionDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }

    public class FilterAfterTableConversionDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");

            // Add a ListObject (table) that includes an auto‑filter by default
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Verify that the table currently has an auto‑filter
            Console.WriteLine("HasAutoFilter before conversion: " + table.HasAutoFilter);

            // Convert the table back to a normal range
            table.ConvertToRange();

            // After conversion the ListObject is removed; attempting to use its Filter method should raise an exception
            try
            {
                // This call is expected to fail because the object is no longer a table
                table.Filter();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Expected exception caught after conversion: " + ex.Message);
            }

            // Optionally, apply a worksheet‑level filter to the same range to show that it still works
            try
            {
                CellArea filterArea = new CellArea
                {
                    StartRow = 0,      // Row 1 (zero‑based)
                    StartColumn = 0,   // Column A
                    EndRow = 3,        // Row 4
                    EndColumn = 1      // Column B
                };
                sheet.Filter(filterArea);
                Console.WriteLine("Worksheet filter applied successfully to the former table range.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error applying worksheet filter: " + ex.Message);
            }

            // Save the workbook (optional, just to complete the lifecycle)
            try
            {
                workbook.Save("FilterAfterTableConversionDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
