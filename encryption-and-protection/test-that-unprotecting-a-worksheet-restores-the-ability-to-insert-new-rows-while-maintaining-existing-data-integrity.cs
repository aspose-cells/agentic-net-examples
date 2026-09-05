// Title: Unprotect an Aspose.Cells worksheet in C# to insert rows while preserving existing cell values
// AI Prompts: Write C# code that protects a worksheet with a password using Aspose.Cells, attempts to insert a row (expecting an exception), then unprotects the sheet and successfully inserts a new row. | Show how to iterate through the first column after unprotecting a worksheet to confirm that original data remains unchanged. | Explain how to catch and handle the exception thrown when inserting rows into a password‑protected worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# unprotect worksheet before inserting rows example | how to catch row insertion error on protected Excel sheet using Aspose.Cells | verify data integrity after unprotecting a password‑protected worksheet in C# | insert blank row between header and data after removing protection with Aspose.Cells | save workbook after unprotecting and modifying rows in C# Aspose.Cells
// Tags: worksheet unprotect Aspose.Cells C# | insert rows after worksheet unprotected Aspose.Cells | password protected Excel sheet Aspose.Cells handling | exception on row insertion in protected worksheet Aspose.Cells | data integrity verification after row insertion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, protects the first worksheet with a password, attempts a row insertion that throws an exception, then unprotects the sheet, inserts a new row, iterates through the column to confirm original data is intact, and finally saves the workbook.
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

            // Populate initial data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row1");
            sheet.Cells["A3"].PutValue("Row2");

            // Protect the worksheet with a password (oldPassword is empty for new protection)
            sheet.Protect(ProtectionType.All, "SecretPwd", string.Empty);

            // Attempt to insert a row while the sheet is protected (should throw)
            bool insertFailed = false;
            try
            {
                // Insert one row after the header (row index 1)
                sheet.Cells.InsertRows(1, 1);
            }
            catch (Exception ex)
            {
                insertFailed = true;
                Console.WriteLine("Expected failure while inserting on protected sheet: " + ex.Message);
            }

            if (!insertFailed)
            {
                Console.WriteLine("Error: Row insertion succeeded despite protection.");
            }

            // Unprotect the worksheet
            sheet.Unprotect("SecretPwd");

            // Insert a new row now that the sheet is unprotected
            sheet.Cells.InsertRows(1, 1); // Inserts a blank row between Header and Row1

            // Verify data integrity after insertion
            Console.WriteLine("Data after unprotected insertion:");
            for (int row = 0; row < 4; row++)
            {
                Console.WriteLine($"Row {row + 1}: {sheet.Cells[row, 0].StringValue}");
            }

            // Save the workbook (optional, demonstrates that the file is valid)
            string outputPath = "UnprotectWorksheetTest.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
