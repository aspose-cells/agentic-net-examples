// Title: Aspose.Cells C# – Verify ListObject conversion to a range removes the table and disables sorting
// Description: Creates a workbook, adds a ListObject (table) over A1:B3, calls ConvertToRange, then confirms the ListObject collection is empty and ListObject.IsTable returns false, proving the table can no longer be sorted.
// Keywords: Aspose.Cells ConvertToRange | ListObject.IsTable false | remove table Aspose.Cells | C# Aspose.Cells table to range | verify table removal Aspose.Cells | disable sorting ListObject | Aspose.Cells ListObject conversion
// Common Searches: Aspose.Cells ConvertToRange example | Check if ListObject still exists after conversion | ListObject.IsTable after ConvertToRange | How to delete a table in Aspose.Cells .NET | Disable sorting on Aspose.Cells table
// Developer Intent: Ensure that converting a ListObject to a normal range eliminates the table so it no longer supports sorting or other table features.
// Use Cases: Add a ListObject to a worksheet, convert it to a range, and verify sheet.ListObjects.Count is zero. | After conversion, query ListObject.IsTable (or catch its absence) to confirm the object is no longer a table. | Save the workbook to persist the removal of table metadata.
// AI Prompts: Generate C# code using Aspose.Cells that adds a table, converts it to a range, and asserts ListObject.IsTable is false after conversion. | Write a unit test in .NET that verifies a ListObject is removed from the worksheet after calling ConvertToRange. | Explain step‑by‑step how to confirm a table no longer supports sorting after converting it to a normal range with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a ListObject (table) over A1:B3, calls ConvertToRange, then confirms the ListObject collection is empty and ListObject.IsTable returns false, proving the table can no longer be sorted.
class VerifyTableConversion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");

            // Add a ListObject (table) covering the data range
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Verify that a ListObject (table) exists before conversion
            Console.WriteLine("Before conversion - ListObject exists: " + (sheet.ListObjects.Count > 0));

            // Convert the table back to a normal range; this removes the ListObject
            table.ConvertToRange();

            // After conversion the ListObject should be removed from the collection
            bool tableExists = sheet.ListObjects.Count > 0;
            Console.WriteLine("After conversion - ListObject exists: " + tableExists);

            // Save the workbook
            string outputPath = "VerifyTableConversion.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
