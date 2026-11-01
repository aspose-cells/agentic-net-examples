// Title: How to unprotect an Excel worksheet and insert rows while preserving data integrity using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unprotect worksheet then insert a row | verify row insertion after removing protection with Aspose.Cells | C# example to test worksheet protection and data shift in Excel using Aspose.Cells | how to insert rows in a protected Excel sheet with Aspose.Cells | check data integrity after unprotecting a sheet in Aspose.Cells .NET
// Tags: worksheet security removal Aspose.Cells | post-unprotect row insertion API Aspose.Cells | worksheet protection handling Aspose.Cells .NET | excel data shift verification Aspose.Cells | save workbook after modifications Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, protects the first worksheet with a password, confirms that row insertion fails while protected, then unprotects the sheet, inserts a new row, validates that existing data shifts correctly and integrity is maintained, and finally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "TestSheet";

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row1");
            sheet.Cells["A3"].PutValue("Row2");

            // Protect the worksheet with a password (oldPassword is not required, pass null)
            sheet.Protect(ProtectionType.All, "pwd123", null);

            // Attempt to insert a row while the sheet is protected (should throw)
            bool insertFailed = false;
            try
            {
                sheet.Cells.InsertRows(2, 1); // Insert before row 2
            }
            catch (Exception ex)
            {
                insertFailed = true;
                Console.WriteLine("Insert failed as expected while protected: " + ex.Message);
            }

            if (!insertFailed)
            {
                Console.WriteLine("Error: Row insertion succeeded despite protection.");
            }

            // Unprotect the worksheet
            sheet.Unprotect("pwd123");

            // Insert a new row now that the sheet is unprotected
            sheet.Cells.InsertRows(2, 1); // Insert before row 2

            // Verify that existing data shifted correctly
            string header = sheet.Cells["A1"].StringValue; // Should be "Header"
            string insertedRow = sheet.Cells["A2"].StringValue; // Should be empty
            string row1 = sheet.Cells["A3"].StringValue; // Should be "Row1"
            string row2 = sheet.Cells["A4"].StringValue; // Should be "Row2"

            Console.WriteLine($"Header: {header}");
            Console.WriteLine($"Inserted Row (expected empty): '{insertedRow}'");
            Console.WriteLine($"Row1 after insert: {row1}");
            Console.WriteLine($"Row2 after insert: {row2}");

            // Simple integrity check
            if (header == "Header" && row1 == "Row1" && row2 == "Row2")
            {
                Console.WriteLine("Data integrity maintained after unprotecting and inserting rows.");
            }
            else
            {
                Console.WriteLine("Data integrity issue detected.");
            }

            // Save the workbook for manual inspection if needed
            string outputPath = "UnprotectTest.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to save workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
