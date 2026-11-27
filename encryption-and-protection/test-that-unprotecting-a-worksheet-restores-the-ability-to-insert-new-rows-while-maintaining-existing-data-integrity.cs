// Title: Unprotect an Excel worksheet and insert rows without losing data using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unprotect worksheet then insert a row | insert rows after removing protection in an Excel file with Aspose.Cells | verify data integrity after unprotecting a protected sheet using Aspose.Cells | C# example for worksheet protection, unprotection and row insertion with Aspose.Cells | how to handle row insertion failure on a protected worksheet in Aspose.Cells
// Tags: worksheet unprotect API Aspose.Cells | insert rows after worksheet unprotect C# | protect worksheet password Aspose.Cells | data integrity validation after row insertion | exception handling for protected sheet insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetUnprotectDemo
{
    // The example creates a workbook, protects the first worksheet with a password, demonstrates that row insertion fails while protected, then unprotects the sheet, inserts a new row, checks that original cell values remain unchanged and the new row is empty, and finally saves the workbook.
    class WorksheetUnprotectTest
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate initial data
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Row1");
                sheet.Cells["A3"].PutValue("Row2");

                // Protect the worksheet with a password (oldPassword not required for a new sheet)
                sheet.Protect(ProtectionType.All, "pwd123", string.Empty);

                // Attempt to insert a row while protected (should fail)
                bool insertFailed = false;
                try
                {
                    sheet.Cells.InsertRows(2, 1);
                }
                catch (Exception ex)
                {
                    insertFailed = true;
                    Console.WriteLine("Insert failed as expected while protected: " + ex.Message);
                }

                if (!insertFailed)
                {
                    Console.WriteLine("Error: Row insertion succeeded while worksheet was protected.");
                }

                // Unprotect the worksheet
                sheet.Unprotect("pwd123");

                // Insert a new row now that the sheet is unprotected
                sheet.Cells.InsertRows(2, 1);

                // Verify data integrity after insertion
                string header = sheet.Cells["A1"].StringValue;
                string newRow = sheet.Cells["A2"].StringValue; // should be empty
                string row1 = sheet.Cells["A3"].StringValue;
                string row2 = sheet.Cells["A4"].StringValue;

                Console.WriteLine("After unprotect and insert:");
                Console.WriteLine($"A1: {header}");
                Console.WriteLine($"A2 (new row): '{newRow}'");
                Console.WriteLine($"A3 (original Row1): {row1}");
                Console.WriteLine($"A4 (original Row2): {row2}");

                if (header == "Header" && string.IsNullOrEmpty(newRow) && row1 == "Row1" && row2 == "Row2")
                {
                    Console.WriteLine("Data integrity maintained after unprotecting and inserting rows.");
                }
                else
                {
                    Console.WriteLine("Data integrity check failed.");
                }

                // Save the workbook
                string outputPath = "WorksheetUnprotectTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception e)
            {
                Console.WriteLine("Unexpected error: " + e.Message);
            }
        }
    }
}
