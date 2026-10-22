// Title: Protect an Excel workbook’s structure and confirm worksheet rename is blocked using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells protect workbook structure C# example | how to block worksheet rename after protecting workbook with Aspose.Cells | C# code to verify rename operation is prevented when workbook structure is locked | Aspose.Cells exception handling for renaming sheet on a protected workbook | save Excel file after applying structure protection with Aspose.Cells
// Tags: Aspose.Cells workbook structure lock C# | worksheet rename restriction Aspose.Cells | structure protection exception handling C# | save protected Excel workbook Aspose.Cells | verify sheet rename blocked .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, names the first worksheet, applies structure protection (without a password), attempts to rename the sheet, catches the expected exception to confirm the rename is blocked, and finally saves the protected workbook to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set initial name for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Original";

            // Protect the workbook structure (prevents adding, deleting, moving, renaming sheets)
            // Provide an empty password to use default protection
            workbook.Protect(ProtectionType.Structure, string.Empty);

            // Attempt to rename the worksheet
            bool renameSucceeded = true;
            try
            {
                sheet.Name = "Renamed";
            }
            catch (Exception ex)
            {
                // Expected: operation is blocked because the structure is protected
                renameSucceeded = false;
                Console.WriteLine("Rename blocked: " + ex.Message);
            }

            // Confirm that the rename operation was blocked
            if (!renameSucceeded || sheet.Name != "Renamed")
            {
                Console.WriteLine("Worksheet rename was blocked as expected.");
            }
            else
            {
                Console.WriteLine("Worksheet rename succeeded unexpectedly.");
            }

            // Save the workbook (optional, demonstrates that the file is still valid)
            string outputPath = "ProtectedWorkbook.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
