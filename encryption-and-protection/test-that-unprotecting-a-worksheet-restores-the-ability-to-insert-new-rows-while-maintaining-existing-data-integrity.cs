// Title: C# – Unprotect an Aspose.Cells worksheet to insert rows while preserving data integrity
// Description: Demonstrates protecting a worksheet with row insertion disabled, catching the expected failure, unprotecting with a password, inserting a new row, verifying that original cells shift correctly, and saving the workbook.
// Keywords: Aspose.Cells protect worksheet C# | unprotect worksheet Aspose.Cells | insert row after unprotect | row insertion disabled protection | verify data integrity Aspose.Cells | C# Excel protection example | Aspose.Cells InsertRow exception
// Common Searches: how to unprotect an Aspose.Cells worksheet in C# | Aspose.Cells insert row after sheet protection | C# protect sheet without allowing row insertion | Aspose.Cells verify cell values after unprotect | example of worksheet protection and unprotect in Aspose.Cells
// Developer Intent: Remove protection from a worksheet so that a new row can be added without altering existing cell values.
// Use Cases: Lock a template sheet to prevent manual row additions, then programmatically unprotect it to insert summary rows before export. | Create a report with a fixed layout, temporarily unprotect the sheet to add calculated rows, and re‑apply protection if needed. | Automated test that confirms InsertRow throws an exception on a protected sheet and succeeds after calling Unprotect with the correct password.
// AI Prompts: Generate C# code using Aspose.Cells that protects a worksheet, blocks row insertion, catches the failure, then unprotects with a password and inserts a new row while keeping existing data intact. | Explain how to handle the exception thrown by InsertRow on a protected sheet and how to validate that original rows shift correctly after unprotecting. | Write a unit test in C# that asserts InsertRow fails on a protected worksheet and passes after calling Unprotect with the correct password.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionTest
{
    // Demonstrates protecting a worksheet with row insertion disabled, catching the expected failure, unprotecting with a password, inserting a new row, verifying that original cells shift correctly, and saving the workbook.
    public class UnprotectInsertRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate initial data in rows 0 to 2
                cells[0, 0].PutValue("Row0");
                cells[0, 1].PutValue(10);
                cells[1, 0].PutValue("Row1");
                cells[1, 1].PutValue(20);
                cells[2, 0].PutValue("Row2");
                cells[2, 1].PutValue(30);

                // Configure protection: disallow inserting rows
                Protection protection = worksheet.Protection;
                protection.AllowInsertingRow = false;
                protection.Password = "pwd123";
                worksheet.Protect(ProtectionType.All, "pwd123", null);

                Console.WriteLine($"Worksheet protected: {worksheet.IsProtected}");

                // Attempt to insert a row while protected (should fail)
                try
                {
                    cells.InsertRow(1);
                    Console.WriteLine("Unexpected: Row inserted while protection disallows it.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Expected failure when inserting row on protected sheet: {ex.Message}");
                }

                // Unprotect the worksheet with the correct password
                worksheet.Unprotect("pwd123");
                Console.WriteLine($"Worksheet protected after unprotect: {worksheet.IsProtected}");

                // Insert a new row at index 1 (between original Row0 and Row1)
                cells.InsertRow(1);
                // Add data to the newly inserted row
                cells[1, 0].PutValue("InsertedRow");
                cells[1, 1].PutValue(99);

                // Verify data integrity after insertion
                Console.WriteLine($"Cell A0: {cells[0, 0].StringValue} (expected 'Row0')");
                Console.WriteLine($"Cell A1: {cells[1, 0].StringValue} (expected 'InsertedRow')");
                Console.WriteLine($"Cell A2: {cells[2, 0].StringValue} (expected 'Row1')");
                Console.WriteLine($"Cell A3: {cells[3, 0].StringValue} (expected 'Row2')");

                // Save the workbook to verify the result
                workbook.Save("UnprotectInsertRowDemo.xlsx");
                Console.WriteLine("Workbook saved as 'UnprotectInsertRowDemo.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnprotectInsertRowDemo.Run();
        }
    }
}
