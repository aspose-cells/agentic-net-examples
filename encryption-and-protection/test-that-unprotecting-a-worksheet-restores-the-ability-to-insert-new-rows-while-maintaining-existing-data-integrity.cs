// Title: Aspose.Cells for .NET – Unprotect a Worksheet and Insert a Row Without Losing Data
// Description: C# example that creates a workbook, protects the first worksheet, demonstrates that row insertion fails while protected, then unprotects with the correct password, inserts a new row, shifts existing values, and saves the file. Shows how to verify data integrity after unprotecting.
// Keywords: Aspose.Cells unprotect worksheet C# | insert row after unprotect Aspose.Cells | worksheet protection row insertion | data integrity after row insert | C# Aspose.Cells protect sheet example | Unprotect InsertRow Aspose.Cells
// Common Searches: how to insert a row in a protected worksheet using Aspose.Cells | Aspose.Cells C# unprotect sheet then add row | verify data shift after inserting row in Aspose.Cells | Aspose.Cells protect sheet prevent row insertion | C# test worksheet unprotect restores insert row
// Developer Intent: Confirm that calling Unprotect on a protected worksheet re‑enables row insertion and that existing cell values remain correctly positioned after the new row is added.
// Use Cases: Automated test to ensure protection blocks row insertion and unprotect restores it without corrupting data. | Workflow that temporarily locks a sheet during calculations, then unlocks it to allow users to add rows before final save. | Validation that a workbook saved after unprotecting can be edited by end users while preserving original content.
// AI Prompts: Write a C# unit test with Aspose.Cells that asserts row insertion throws an exception on a protected sheet and succeeds after Unprotect with the correct password. | Provide C# code that protects a worksheet, attempts to insert a row, catches the failure, then unprotects, inserts a row, and checks that previous cell values have shifted as expected. | Explain how Aspose.Cells ProtectionType controls allowed actions and how to temporarily lift row‑insertion restrictions using Unprotect.

using System;
using Aspose.Cells;

// C# example that creates a workbook, protects the first worksheet, demonstrates that row insertion fails while protected, then unprotects with the correct password, inserts a new row, shifts existing values, and saves the file. Shows how to verify data integrity after unprotecting.
class TestWorksheetUnprotectInsertRow
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate initial data in rows 0, 1 and 2
        cells[0, 0].PutValue("Row0");
        cells[1, 0].PutValue("Row1");
        cells[2, 0].PutValue("Row2");

        // Protect the worksheet with a password (insertion of rows is not allowed by default)
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Attempt to insert a row while the sheet is protected – this should fail
        try
        {
            cells.InsertRow(1);
            Console.WriteLine("Unexpected: Row inserted while worksheet is protected.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Expected failure inserting row while protected: " + ex.Message);
        }

        // Unprotect the worksheet using the correct password
        sheet.Unprotect("pwd123");

        // Insert a new row at index 1 (original rows shift down)
        cells.InsertRow(1);
        cells[1, 0].PutValue("InsertedRow");

        // Verify that existing data remains intact and has shifted correctly
        Console.WriteLine("Cell A0 (should be Row0): " + cells[0, 0].StringValue);
        Console.WriteLine("Cell A1 (should be InsertedRow): " + cells[1, 0].StringValue);
        Console.WriteLine("Cell A2 (should be Row1): " + cells[2, 0].StringValue);
        Console.WriteLine("Cell A3 (should be Row2): " + cells[3, 0].StringValue);

        // Save the workbook to verify the result
        workbook.Save("UnprotectInsertRowDemo.xlsx");
    }
}
