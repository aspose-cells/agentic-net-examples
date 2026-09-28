// Title: How to catch ArgumentOutOfRangeException when adding a slicer to a non‑existent table column in Aspose.Cells for .NET
// AI Prompts: Write C# code that validates a ListObject column index before creating a slicer and wraps the operation in a try‑catch block for Aspose.Cells. | Show an example that catches ArgumentOutOfRangeException when calling Worksheet.Slicers.Add with an out‑of‑range column in Aspose.Cells. | Demonstrate how to log the error and continue processing after a slicer addition fails because the target table column is missing in Aspose.Cells.
// Common Searches: aspnet add slicer to table column index out of range exception handling | c# Aspose.Cells slicer creation fails when column does not exist | how to validate slicer column index before adding in Aspose.Cells workbook | catch ArgumentOutOfRangeException for slicer.Add in Aspose.Cells .NET
// Tags: Aspose.Cells slicer column existence check | C# try‑catch for slicer addition errors | ListObject column range validation Aspose.Cells | Workbook save after slicer failure Aspose.Cells | Exception handling pattern for Aspose.Cells slicers

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace AsposeCellsSlicerErrorHandling
{
    // The example creates a workbook with a two‑column table, then attempts to add a slicer for a column index that does not exist. The slicer addition is enclosed in a try‑catch block that specifically handles ArgumentOutOfRangeException and a generic Exception, logs appropriate messages, and finally saves the workbook regardless of success.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();                     // create rule
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for a table (2 columns)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue("Fruit");
            worksheet.Cells["A3"].PutValue("Carrot");
            worksheet.Cells["B3"].PutValue("Vegetable");

            // Add a ListObject (table) covering the data range
            int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Attempt to add a slicer for a column that does NOT exist (e.g., index 5)
            try
            {
                // This will throw if the column index is out of range
                ListColumn nonExistentColumn = table.ListColumns[5]; // zero‑based index
                // Add slicer using the valid overload that accepts ListColumn and destination cell name
                SlicerCollection slicers = worksheet.Slicers;        // property rule
                slicers.Add(table, nonExistentColumn, "D5");        // add rule
                Console.WriteLine("Slicer added successfully.");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Handle the specific case where the column index is invalid
                Console.WriteLine("Error: The specified table column does not exist.");
                Console.WriteLine($"Details: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected errors
                Console.WriteLine("An unexpected error occurred while adding the slicer.");
                Console.WriteLine($"Details: {ex.Message}");
            }

            // Save the workbook (even if slicer addition failed)
            workbook.Save("SlicerErrorHandlingOutput.xlsx");        // save rule
        }
    }
}
