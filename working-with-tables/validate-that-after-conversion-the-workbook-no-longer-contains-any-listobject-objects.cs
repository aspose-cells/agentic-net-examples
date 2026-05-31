using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will become a table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("John");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Mary");

        // Add a ListObject (table) to the worksheet
        int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Convert the ListObject to a normal range
        table.ConvertToRange();

        // Validate that the worksheet no longer contains any ListObject objects
        if (sheet.ListObjects.Count == 0)
        {
            Console.WriteLine("Validation passed: No ListObject objects remain after conversion.");
        }
        else
        {
            Console.WriteLine($"Validation failed: {sheet.ListObjects.Count} ListObject(s) still present.");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ValidatedWorkbook.xlsx");
    }
}