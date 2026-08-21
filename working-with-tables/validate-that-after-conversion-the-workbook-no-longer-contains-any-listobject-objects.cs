// Title: Remove all ListObject tables after converting them to ranges with Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a ListObject (table) on A1:B3, converts every ListObject in every worksheet to a normal range using ConvertToRange, verifies that no tables remain, and saves the file.
// Keywords: Aspose.Cells ListObject conversion | ConvertToRange C# | delete tables Aspose.Cells | check for ListObjects after conversion | .NET spreadsheet table removal | Aspose.Cells workbook cleanup
// Common Searches: how to delete ListObject after ConvertToRange Aspose.Cells | verify no tables exist in workbook C# Aspose.Cells | convert Excel tables to ranges programmatically | Aspose.Cells remove all ListObjects from workbook | C# sample for ListObject cleanup in Aspose.Cells
// Developer Intent: Ensure that every ListObject in a workbook has been transformed into a regular range and that the workbook no longer contains any table objects.
// Use Cases: Prepare a workbook for CSV export where tables cause formatting issues. | Apply cell‑level styling that is unsupported on table objects. | Automated testing to confirm ListObject cleanup after batch processing.
// AI Prompts: Generate a C# function that returns true if any ListObjects are still present after calling ConvertToRange on all worksheets using Aspose.Cells. | Write a unit test in NUnit that asserts no ListObjects remain after converting all tables to ranges in an Aspose.Cells workbook. | Suggest an alternative method to remove ListObjects without using ConvertToRange, leveraging other Aspose.Cells APIs.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This C# example creates a workbook, adds a ListObject (table) on A1:B3, converts every ListObject in every worksheet to a normal range using ConvertToRange, verifies that no tables remain, and saves the file.
class ValidateListObjectsRemoval
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will become a table
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("John");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Jane");

        // Add a ListObject (table) to the worksheet
        int loIndex = sheet.ListObjects.Add("A1", "B3", true);
        ListObject listObject = sheet.ListObjects[loIndex];

        // Convert every ListObject in every worksheet to a normal range
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Iterate backwards because ConvertToRange removes the ListObject from the collection
            for (int i = ws.ListObjects.Count - 1; i >= 0; i--)
            {
                ws.ListObjects[i].ConvertToRange();
            }
        }

        // Verify that no ListObjects remain in the workbook
        bool anyListObjects = false;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.ListObjects.Count > 0)
            {
                anyListObjects = true;
                break;
            }
        }

        Console.WriteLine(anyListObjects
            ? "ListObjects still exist after conversion."
            : "All ListObjects have been removed.");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("ValidatedWorkbook.xlsx");
    }
}
