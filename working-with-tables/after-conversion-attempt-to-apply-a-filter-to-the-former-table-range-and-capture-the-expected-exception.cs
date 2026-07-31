// Title: Handling Expected Exceptions When Filtering After Converting an Aspose.Cells Table to a Range (C#)
// Description: C# example that creates a workbook, adds a ListObject with an auto‑filter, converts the table to a normal range, then attempts Worksheet.Filter and ListObject.Filter on the former table area. The code catches the exceptions thrown by these invalid filter operations and saves the workbook.
// Keywords: Aspose.Cells ConvertToRange | C# ListObject filter exception | Worksheet.Filter error after table conversion | Aspose.Cells handling filter exceptions | convert table to range Aspose.Cells
// Common Searches: Aspose.Cells exception when applying Worksheet.Filter after ConvertToRange | ListObject.Filter after table conversion error C# | how to catch filter errors in Aspose.Cells | cannot filter a range that was a table Aspose.Cells
// Developer Intent: Demonstrate how to detect and handle the exceptions raised when applying filters to a range that was previously a table after using ConvertToRange in Aspose.Cells.
// Use Cases: Validate that Worksheet.Filter is not applicable to a converted table range and capture the resulting exception. | Show that invoking ListObject.Filter after ConvertToRange throws an error, enabling defensive coding. | Ensure workbook processing continues by saving the file even when filter operations fail.
// AI Prompts: Generate C# code with Aspose.Cells that converts a ListObject to a range, tries Worksheet.Filter on the original area, and logs any exceptions. | Explain why Worksheet.Filter throws an exception after a table is converted to a range in Aspose.Cells and propose a check to avoid the call. | Write a unit test in C# that asserts the expected exceptions are thrown when filtering a converted table range using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a ListObject with an auto‑filter, converts the table to a normal range, then attempts Worksheet.Filter and ListObject.Filter on the former table area. The code catches the exceptions thrown by these invalid filter operations and saves the workbook.
    public class ConvertTableToRangeAndFilterDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including a header row)
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

            // Convert the table to a normal range
            table.ConvertToRange();

            // Define the cell area that previously represented the table
            CellArea formerTableArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 3,        // Row 4
                EndColumn = 1      // Column B
            };

            // Attempt to apply a worksheet filter to the former table range
            try
            {
                sheet.Filter(formerTableArea);
                Console.WriteLine("Worksheet.Filter applied successfully to the former table range.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Expected exception from Worksheet.Filter: {ex.Message}");
            }

            // Attempt to use the obsolete ListObject.Filter method after conversion
            try
            {
                // This should raise an exception because the ListObject no longer has an AutoFilter
                AutoFilter autoFilter = table.Filter(); // Obsolete but kept for demonstration
                // If no exception, apply a simple filter (won't be reached in normal scenario)
                autoFilter.Filter(0, "1");
                Console.WriteLine("ListObject.Filter applied (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Expected exception from ListObject.Filter: {ex.Message}");
            }

            // Save the workbook (optional, just to complete the lifecycle)
            try
            {
                workbook.Save("ConvertTableToRangeAndFilterDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ConvertTableToRangeAndFilterDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
